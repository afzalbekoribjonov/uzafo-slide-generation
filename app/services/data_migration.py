from __future__ import annotations

import logging
from datetime import datetime, timezone
from typing import Any

from motor.motor_asyncio import AsyncIOMotorDatabase

from app.db.mongo import Mongo

logger = logging.getLogger(__name__)


def utcnow() -> datetime:
    return datetime.now(timezone.utc)


class MigrationConfigError(RuntimeError):
    pass


class MigrationBlockedError(RuntimeError):
    pass


class LegacyMongoToCurrentDbMigrationService:
    FLAG_ID = 'legacy_mongodb_to_current_db'

    def __init__(
        self,
        *,
        current_db: AsyncIOMotorDatabase,
        legacy_mongodb_uri: str | None,
        legacy_mongodb_db: str | None,
    ) -> None:
        self.current_db = current_db
        self.legacy_mongodb_uri = (legacy_mongodb_uri or '').strip() or None
        self.legacy_mongodb_db = (legacy_mongodb_db or '').strip() or None
        self.flags_collection = current_db.system_flags

    async def get_state(self) -> dict[str, Any] | None:
        return await self.flags_collection.find_one({'_id': self.FLAG_ID})

    async def run(self, *, admin_id: int) -> dict[str, Any]:
        if not self.legacy_mongodb_uri:
            raise MigrationConfigError('LEGACY_MONGODB_URI topilmadi.')

        state = await self.get_state()
        if state and state.get('status') == 'done':
            raise MigrationBlockedError('Bu migratsiya allaqachon muvaffaqiyatli bajarilgan.')
        if state and state.get('status') == 'running':
            raise MigrationBlockedError('Migratsiya hozir ham ishlamoqda. Avval uning yakunlanishini kuting.')

        await self._mark_running(admin_id=admin_id)
        legacy_db_name = self.legacy_mongodb_db or self.current_db.name
        legacy_mongo = Mongo(self.legacy_mongodb_uri, legacy_db_name)

        try:
            await legacy_mongo.db.command('ping')
            summary = {
                'legacy_db_name': legacy_db_name,
                'users': await self._migrate_users(legacy_mongo.db),
                'referrals': await self._migrate_referrals(legacy_mongo.db),
                'mandatory_channels': await self._migrate_channels(legacy_mongo.db),
                'generations': await self._migrate_generations(legacy_mongo.db),
            }
            await self._mark_done(admin_id=admin_id, summary=summary)
            return summary
        except Exception as exc:
            logger.exception('Legacy MongoDB migratsiyasi muvaffaqiyatsiz tugadi.')
            await self._mark_failed(admin_id=admin_id, error=str(exc))
            raise
        finally:
            await legacy_mongo.close()

    async def _mark_running(self, *, admin_id: int) -> None:
        now = utcnow()
        await self.flags_collection.update_one(
            {'_id': self.FLAG_ID},
            {
                '$set': {
                    'status': 'running',
                    'started_by': admin_id,
                    'started_at': now,
                    'finished_at': None,
                    'last_error': None,
                },
                '$setOnInsert': {
                    'created_at': now,
                },
            },
            upsert=True,
        )

    async def _mark_done(self, *, admin_id: int, summary: dict[str, Any]) -> None:
        await self.flags_collection.update_one(
            {'_id': self.FLAG_ID},
            {
                '$set': {
                    'status': 'done',
                    'finished_at': utcnow(),
                    'finished_by': admin_id,
                    'summary': summary,
                    'last_error': None,
                }
            },
        )

    async def _mark_failed(self, *, admin_id: int, error: str) -> None:
        await self.flags_collection.update_one(
            {'_id': self.FLAG_ID},
            {
                '$set': {
                    'status': 'failed',
                    'finished_at': utcnow(),
                    'finished_by': admin_id,
                    'last_error': error,
                }
            },
            upsert=True,
        )

    async def _migrate_users(self, legacy_db: AsyncIOMotorDatabase) -> dict[str, int]:
        stats = {'inserted': 0, 'merged': 0, 'skipped': 0}
        async for legacy_user in legacy_db.users.find({}):
            telegram_id = self._safe_int(legacy_user.get('telegram_id'))
            if telegram_id <= 0:
                stats['skipped'] += 1
                continue

            current_user = await self.current_db.users.find_one({'telegram_id': telegram_id})
            if current_user and current_user.get('legacy_mongodb_user_imported'):
                stats['skipped'] += 1
                continue

            source_id = str(legacy_user.get('_id'))
            imported_at = utcnow()
            if not current_user:
                payload = dict(legacy_user)
                payload.pop('_id', None)
                payload['legacy_mongodb_user_imported'] = True
                payload['legacy_mongodb_user_imported_at'] = imported_at
                payload['legacy_mongodb_user_source_id'] = source_id
                await self.current_db.users.update_one(
                    {'telegram_id': telegram_id},
                    {'$setOnInsert': payload},
                    upsert=True,
                )
                stats['inserted'] += 1
                continue

            merged = self._merge_user_documents(current=current_user, legacy=legacy_user)
            merged['legacy_mongodb_user_imported'] = True
            merged['legacy_mongodb_user_imported_at'] = imported_at
            merged['legacy_mongodb_user_source_id'] = source_id
            await self.current_db.users.update_one({'_id': current_user['_id']}, {'$set': merged})
            stats['merged'] += 1
        return stats

    async def _migrate_referrals(self, legacy_db: AsyncIOMotorDatabase) -> dict[str, int]:
        stats = {'inserted': 0, 'updated': 0, 'skipped': 0}
        async for legacy_referral in legacy_db.referrals.find({}):
            invitee_id = self._safe_int(legacy_referral.get('invitee_id'))
            if invitee_id <= 0:
                stats['skipped'] += 1
                continue

            current_referral = await self.current_db.referrals.find_one({'invitee_id': invitee_id})
            merged = self._merge_referral_documents(current=current_referral, legacy=legacy_referral)
            if current_referral:
                await self.current_db.referrals.update_one({'_id': current_referral['_id']}, {'$set': merged})
                stats['updated'] += 1
            else:
                await self.current_db.referrals.insert_one(merged)
                stats['inserted'] += 1
        return stats

    async def _migrate_channels(self, legacy_db: AsyncIOMotorDatabase) -> dict[str, int]:
        stats = {'inserted': 0, 'updated': 0, 'skipped': 0}
        async for legacy_channel in legacy_db.mandatory_channels.find({}):
            chat_id = self._safe_int(legacy_channel.get('chat_id'))
            if chat_id == 0:
                stats['skipped'] += 1
                continue

            current_channel = await self.current_db.mandatory_channels.find_one({'chat_id': chat_id})
            merged = self._merge_channel_documents(current=current_channel, legacy=legacy_channel)
            if current_channel:
                await self.current_db.mandatory_channels.update_one({'_id': current_channel['_id']}, {'$set': merged})
                stats['updated'] += 1
            else:
                await self.current_db.mandatory_channels.insert_one(merged)
                stats['inserted'] += 1
        return stats

    async def _migrate_generations(self, legacy_db: AsyncIOMotorDatabase) -> dict[str, int]:
        stats = {'inserted': 0, 'skipped': 0}
        async for legacy_generation in legacy_db.generations.find({}):
            source_id = str(legacy_generation.get('_id'))
            if not source_id:
                stats['skipped'] += 1
                continue

            exists = await self.current_db.generations.find_one({'legacy_source_id': source_id}, projection={'_id': 1})
            if exists:
                stats['skipped'] += 1
                continue

            payload = dict(legacy_generation)
            payload.pop('_id', None)
            payload['legacy_source'] = 'legacy_mongodb'
            payload['legacy_source_id'] = source_id
            payload['legacy_mongodb_generation_imported_at'] = utcnow()

            status = str(payload.get('status') or '').strip().lower()
            if status in {'queued', 'processing'}:
                payload['status'] = 'failed'
                payload['finished_at'] = self._max_datetime(payload.get('finished_at'), utcnow()) or utcnow()
                original_error = str(payload.get('error') or '').strip()
                payload['error'] = (
                    f'{original_error} Imported from legacy DB while still active.'
                    if original_error
                    else 'Imported from legacy DB while still active.'
                )

            await self.current_db.generations.insert_one(payload)
            stats['inserted'] += 1
        return stats

    def _merge_user_documents(self, *, current: dict[str, Any], legacy: dict[str, Any]) -> dict[str, Any]:
        bot_access_blocked = self._truthy(current.get('bot_access_blocked')) or self._truthy(legacy.get('bot_access_blocked'))
        generation_access_blocked = self._truthy(current.get('generation_access_blocked')) or self._truthy(legacy.get('generation_access_blocked'))
        created_at = self._min_datetime(current.get('created_at'), legacy.get('created_at')) or utcnow()
        last_active_at = self._max_datetime(current.get('last_active_at'), legacy.get('last_active_at')) or utcnow()

        return {
            'telegram_id': self._safe_int(current.get('telegram_id') or legacy.get('telegram_id')),
            'full_name': current.get('full_name') or legacy.get('full_name') or 'Unknown',
            'username': current.get('username') or legacy.get('username'),
            'is_admin': self._truthy(current.get('is_admin')) or self._truthy(legacy.get('is_admin')),
            'subscription_verified': self._truthy(current.get('subscription_verified')) or self._truthy(legacy.get('subscription_verified')),
            'invited_by': current.get('invited_by') if current.get('invited_by') is not None else legacy.get('invited_by'),
            'referral_count': self._safe_int(current.get('referral_count')) + self._safe_int(legacy.get('referral_count')),
            'referral_credits': self._safe_int(current.get('referral_credits')) + self._safe_int(legacy.get('referral_credits')),
            'bonus_generation_credits': self._safe_int(current.get('bonus_generation_credits')) + self._safe_int(legacy.get('bonus_generation_credits')),
            'manual_generation_total_added': self._safe_int(current.get('manual_generation_total_added')) + self._safe_int(legacy.get('manual_generation_total_added')),
            'free_generation_used': self._truthy(current.get('free_generation_used')) or self._truthy(legacy.get('free_generation_used')),
            'generation_unlimited': self._truthy(current.get('generation_unlimited')) or self._truthy(legacy.get('generation_unlimited')),
            'generation_access_blocked': generation_access_blocked,
            'bot_access_blocked': bot_access_blocked,
            'generated_count': self._safe_int(current.get('generated_count')) + self._safe_int(legacy.get('generated_count')),
            'successful_generations': self._safe_int(current.get('successful_generations')) + self._safe_int(legacy.get('successful_generations')),
            'status': 'blocked' if bot_access_blocked else (current.get('status') or legacy.get('status') or 'active'),
            'created_at': created_at,
            'last_active_at': last_active_at,
        }

    def _merge_referral_documents(self, *, current: dict[str, Any] | None, legacy: dict[str, Any]) -> dict[str, Any]:
        if current is None:
            payload = dict(legacy)
            payload.pop('_id', None)
            payload['legacy_mongodb_referral_source_id'] = str(legacy.get('_id'))
            payload['legacy_mongodb_referral_imported_at'] = utcnow()
            return payload

        return {
            'inviter_id': current.get('inviter_id') if current.get('inviter_id') is not None else legacy.get('inviter_id'),
            'invitee_id': self._safe_int(current.get('invitee_id') or legacy.get('invitee_id')),
            'start_detected': self._truthy(current.get('start_detected')) or self._truthy(legacy.get('start_detected')),
            'subscription_verified': self._truthy(current.get('subscription_verified')) or self._truthy(legacy.get('subscription_verified')),
            'counted': self._truthy(current.get('counted')) or self._truthy(legacy.get('counted')),
            'created_at': self._min_datetime(current.get('created_at'), legacy.get('created_at')) or utcnow(),
            'counted_at': self._max_datetime(current.get('counted_at'), legacy.get('counted_at')),
            'legacy_mongodb_referral_source_id': str(legacy.get('_id')),
            'legacy_mongodb_referral_imported_at': utcnow(),
        }

    def _merge_channel_documents(self, *, current: dict[str, Any] | None, legacy: dict[str, Any]) -> dict[str, Any]:
        if current is None:
            payload = dict(legacy)
            payload.pop('_id', None)
            payload['legacy_mongodb_channel_source_id'] = str(legacy.get('_id'))
            payload['legacy_mongodb_channel_imported_at'] = utcnow()
            return payload

        return {
            'chat_id': self._safe_int(current.get('chat_id') or legacy.get('chat_id')),
            'title': current.get('title') or legacy.get('title'),
            'username': current.get('username') or legacy.get('username'),
            'invite_link': current.get('invite_link') or legacy.get('invite_link'),
            'is_active': self._truthy(current.get('is_active')) or self._truthy(legacy.get('is_active')),
            'created_at': self._min_datetime(current.get('created_at'), legacy.get('created_at')) or utcnow(),
            'legacy_mongodb_channel_source_id': str(legacy.get('_id')),
            'legacy_mongodb_channel_imported_at': utcnow(),
        }

    @staticmethod
    def _safe_int(value: Any) -> int:
        try:
            return int(value or 0)
        except Exception:
            return 0

    @staticmethod
    def _truthy(value: Any) -> bool:
        return bool(value)

    @staticmethod
    def _min_datetime(first: Any, second: Any) -> datetime | None:
        values = [item for item in (first, second) if isinstance(item, datetime)]
        if not values:
            return None
        return min(values)

    @staticmethod
    def _max_datetime(first: Any, second: Any) -> datetime | None:
        values = [item for item in (first, second) if isinstance(item, datetime)]
        if not values:
            return None
        return max(values)
