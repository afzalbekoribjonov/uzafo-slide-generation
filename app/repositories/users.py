from __future__ import annotations

from datetime import datetime, timedelta, timezone
from typing import Any, AsyncIterator

from motor.motor_asyncio import AsyncIOMotorCollection
from pymongo import DESCENDING, ReturnDocument


def utcnow() -> datetime:
    return datetime.now(timezone.utc)


class UsersRepository:
    def __init__(self, collection: AsyncIOMotorCollection) -> None:
        self.collection = collection

    async def get_by_telegram_id(self, telegram_id: int) -> dict[str, Any] | None:
        return await self.collection.find_one({'telegram_id': telegram_id})

    async def create(
        self,
        *,
        telegram_id: int,
        full_name: str,
        username: str | None,
        invited_by: int | None,
        is_admin: bool = False,
    ) -> dict[str, Any]:
        document = {
            'telegram_id': telegram_id,
            'full_name': full_name,
            'username': username,
            'is_admin': is_admin,
            'subscription_verified': False,
            'invited_by': invited_by,
            'referral_count': 0,
            'referral_credits': 0,
            'bonus_generation_credits': 0,
            'manual_generation_total_added': 0,
            'free_generation_used': False,
            'generation_unlimited': False,
            'generation_access_blocked': False,
            'bot_access_blocked': False,
            'generated_count': 0,
            'successful_generations': 0,
            'status': 'active',
            'created_at': utcnow(),
            'last_active_at': utcnow(),
        }
        await self.collection.insert_one(document)
        return document

    async def touch(self, telegram_id: int, *, full_name: str, username: str | None, is_admin: bool | None = None) -> None:
        set_data: dict[str, Any] = {
            'full_name': full_name,
            'username': username,
            'last_active_at': utcnow(),
        }
        if is_admin is not None:
            set_data['is_admin'] = bool(is_admin)
        await self.collection.update_one(
            {'telegram_id': telegram_id},
            {'$set': set_data},
        )

    async def sync_admin_flag(self, telegram_id: int, is_admin: bool) -> None:
        await self.collection.update_one(
            {'telegram_id': telegram_id},
            {'$set': {'is_admin': bool(is_admin)}},
        )

    async def increment_referral_reward(self, telegram_id: int, credits: int = 1) -> None:
        await self.collection.update_one(
            {'telegram_id': telegram_id},
            {
                '$inc': {
                    'referral_count': 1,
                    'referral_credits': credits,
                },
                '$set': {
                    'last_active_at': utcnow(),
                },
            },
        )

    async def mark_subscription_verified(self, telegram_id: int) -> None:
        await self.collection.update_one(
            {'telegram_id': telegram_id},
            {'$set': {'subscription_verified': True, 'last_active_at': utcnow()}},
        )

    async def set_subscription_verified(self, telegram_id: int, value: bool) -> None:
        await self.collection.update_one(
            {'telegram_id': telegram_id},
            {'$set': {'subscription_verified': bool(value)}},
        )

    async def consume_generation_credit(self, telegram_id: int) -> str | None:
        blocked_user = await self.collection.find_one(
            {
                'telegram_id': telegram_id,
                '$or': [
                    {'generation_access_blocked': True},
                    {'bot_access_blocked': True},
                ],
            },
            projection={'telegram_id': 1},
        )
        if blocked_user:
            return None

        unlimited_result = await self.collection.update_one(
            {
                'telegram_id': telegram_id,
                'generation_unlimited': True,
            },
            {
                '$set': {
                    'last_active_at': utcnow(),
                },
                '$inc': {
                    'generated_count': 1,
                },
            },
        )
        if unlimited_result.modified_count:
            return 'unlimited'

        bonus_result = await self.collection.update_one(
            {
                'telegram_id': telegram_id,
                'bonus_generation_credits': {'$gt': 0},
            },
            {
                '$set': {
                    'last_active_at': utcnow(),
                },
                '$inc': {
                    'bonus_generation_credits': -1,
                    'generated_count': 1,
                },
            },
        )
        if bonus_result.modified_count:
            return 'bonus'

        free_result = await self.collection.update_one(
            {
                'telegram_id': telegram_id,
                'free_generation_used': {'$ne': True},
            },
            {
                '$set': {
                    'free_generation_used': True,
                    'last_active_at': utcnow(),
                },
                '$inc': {
                    'generated_count': 1,
                },
            },
        )
        if free_result.modified_count:
            return 'free'

        referral_result = await self.collection.update_one(
            {
                'telegram_id': telegram_id,
                'referral_credits': {'$gt': 0},
            },
            {
                '$set': {
                    'last_active_at': utcnow(),
                },
                '$inc': {
                    'referral_credits': -1,
                    'generated_count': 1,
                },
            },
        )
        if referral_result.modified_count:
            return 'referral'

        return None

    async def restore_generation_credit(self, telegram_id: int, consumed_from: str) -> None:
        if consumed_from == 'referral':
            await self.collection.update_one(
                {'telegram_id': telegram_id},
                {
                    '$inc': {
                        'referral_credits': 1,
                        'generated_count': -1,
                    },
                    '$set': {
                        'last_active_at': utcnow(),
                    },
                },
            )
            return

        if consumed_from == 'bonus':
            await self.collection.update_one(
                {'telegram_id': telegram_id},
                {
                    '$inc': {
                        'bonus_generation_credits': 1,
                        'generated_count': -1,
                    },
                    '$set': {
                        'last_active_at': utcnow(),
                    },
                },
            )
            return

        if consumed_from == 'unlimited':
            await self.collection.update_one(
                {'telegram_id': telegram_id},
                {
                    '$inc': {
                        'generated_count': -1,
                    },
                    '$set': {
                        'last_active_at': utcnow(),
                    },
                },
            )
            return

        await self.collection.update_one(
            {'telegram_id': telegram_id},
            {
                '$set': {
                    'free_generation_used': False,
                    'last_active_at': utcnow(),
                },
                '$inc': {
                    'generated_count': -1,
                },
            },
        )

    async def increment_successful_generation(self, telegram_id: int) -> None:
        await self.collection.update_one(
            {'telegram_id': telegram_id},
            {
                '$inc': {'successful_generations': 1},
                '$set': {'last_active_at': utcnow()},
            },
        )

    async def set_bot_access_blocked(self, telegram_id: int, value: bool) -> dict[str, Any] | None:
        return await self.collection.find_one_and_update(
            {'telegram_id': telegram_id},
            {
                '$set': {
                    'bot_access_blocked': bool(value),
                    'status': 'blocked' if value else 'active',
                    'last_active_at': utcnow(),
                }
            },
            return_document=ReturnDocument.AFTER,
        )

    async def set_generation_access_blocked(self, telegram_id: int, value: bool) -> dict[str, Any] | None:
        return await self.collection.find_one_and_update(
            {'telegram_id': telegram_id},
            {
                '$set': {
                    'generation_access_blocked': bool(value),
                    'last_active_at': utcnow(),
                }
            },
            return_document=ReturnDocument.AFTER,
        )

    async def set_generation_unlimited(self, telegram_id: int, value: bool) -> dict[str, Any] | None:
        return await self.collection.find_one_and_update(
            {'telegram_id': telegram_id},
            {
                '$set': {
                    'generation_unlimited': bool(value),
                    'last_active_at': utcnow(),
                }
            },
            return_document=ReturnDocument.AFTER,
        )

    async def adjust_bonus_generation_credits(self, telegram_id: int, delta: int) -> dict[str, Any] | None:
        user = await self.get_by_telegram_id(telegram_id)
        if not user:
            return None

        current = int(user.get('bonus_generation_credits', 0) or 0)
        new_value = max(0, current + int(delta))
        manual_total_added = int(user.get('manual_generation_total_added', 0) or 0)
        if delta > 0:
            manual_total_added += int(delta)

        return await self.collection.find_one_and_update(
            {'telegram_id': telegram_id},
            {
                '$set': {
                    'bonus_generation_credits': new_value,
                    'manual_generation_total_added': manual_total_added,
                    'last_active_at': utcnow(),
                }
            },
            return_document=ReturnDocument.AFTER,
        )

    async def search_users(self, query: str, limit: int = 10) -> list[dict[str, Any]]:
        query = (query or '').strip()
        if not query:
            return []

        filters: list[dict[str, Any]] = []
        if query.lstrip('@').isdigit():
            filters.append({'telegram_id': int(query.lstrip('@'))})

        username_value = query.lstrip('@')
        filters.append({'username': {'$regex': username_value, '$options': 'i'}})
        filters.append({'full_name': {'$regex': query, '$options': 'i'}})

        cursor = self.collection.find({'$or': filters}).sort('last_active_at', DESCENDING).limit(limit)
        return await cursor.to_list(length=limit)

    async def count_matching(self, query: dict[str, Any] | None = None) -> int:
        return await self.collection.count_documents(query or {})

    async def list_matching(
        self,
        query: dict[str, Any] | None = None,
        *,
        limit: int = 20,
        sort: list[tuple[str, int]] | None = None,
    ) -> list[dict[str, Any]]:
        cursor = self.collection.find(query or {})
        if sort:
            cursor = cursor.sort(sort)
        cursor = cursor.limit(limit)
        return await cursor.to_list(length=limit)

    async def iterate_matching(
        self,
        query: dict[str, Any] | None = None,
        *,
        sort: list[tuple[str, int]] | None = None,
        batch_size: int = 200,
    ) -> AsyncIterator[dict[str, Any]]:
        cursor = self.collection.find(query or {})
        if sort:
            cursor = cursor.sort(sort)
        cursor = cursor.batch_size(batch_size)
        async for document in cursor:
            yield document

    async def sum_field(self, field_name: str, query: dict[str, Any] | None = None) -> int:
        pipeline = []
        if query:
            pipeline.append({'$match': query})
        pipeline.extend(
            [
                {
                    '$group': {
                        '_id': None,
                        'total': {'$sum': {'$ifNull': [f'${field_name}', 0]}},
                    }
                }
            ]
        )
        items = await self.collection.aggregate(pipeline).to_list(length=1)
        if not items:
            return 0
        return int(items[0].get('total', 0) or 0)

    async def list_top_referrers(self, limit: int = 10) -> list[dict[str, Any]]:
        cursor = (
            self.collection.find({'referral_count': {'$gt': 0}})
            .sort([('referral_count', DESCENDING), ('last_active_at', DESCENDING)])
            .limit(limit)
        )
        return await cursor.to_list(length=limit)

    async def list_top_generators(self, limit: int = 10) -> list[dict[str, Any]]:
        cursor = (
            self.collection.find({'generated_count': {'$gt': 0}})
            .sort([('successful_generations', DESCENDING), ('generated_count', DESCENDING), ('last_active_at', DESCENDING)])
            .limit(limit)
        )
        return await cursor.to_list(length=limit)

    @staticmethod
    def recent_time(days: int) -> datetime:
        return utcnow() - timedelta(days=days)
