from __future__ import annotations

import asyncio
import csv
import re
import tempfile
from datetime import timedelta
from pathlib import Path
from typing import Any

from aiogram import Bot
from aiogram.enums import ParseMode
from aiogram.types import Document, InlineKeyboardButton, InlineKeyboardMarkup, Message
from openpyxl import Workbook
from openpyxl.styles import Font
from pymongo import DESCENDING

from app.callbacks.admin import PublicPostCallback
from app.repositories.channels import ChannelsRepository
from app.repositories.generations import GenerationsRepository
from app.repositories.users import UsersRepository, utcnow
from app.services.generations import GenerationAccessService


class AdminService:
    AUDIENCE_FILTERS: dict[str, str] = {
        'all': '👥 Barcha foydalanuvchilar',
        'subscribed': '✅ Obuna bo‘lganlar',
        'unsubscribed': '⚠️ Obuna bo‘lmaganlar',
        'active_7d': '🔥 So‘nggi 7 kun faollar',
        'inactive_30d': '🕯 30 kundan beri nofaol',
        'referrers': '🏆 Taklif qilganlar',
        'no_credits': '🎟 Limiti tugaganlar',
        'generation_blocked': '🚫 Generation bloklanganlar',
        'bot_blocked': '⛔ Bot kirishi bloklanganlar',
    }

    PUBLIC_BUTTON_ACTIONS: dict[str, str] = {
        'main': 'Asosiy menyu',
        'status': 'Mening holatim',
        'invite': 'Taklif qilish',
        'help': 'Yordam',
        'contact': 'Aloqa',
    }

    def __init__(
        self,
        *,
        users_repo: UsersRepository,
        channels_repo: ChannelsRepository,
        generations_repo: GenerationsRepository,
        generation_access_service: GenerationAccessService,
        bot_username: str,
    ) -> None:
        self.users_repo = users_repo
        self.channels_repo = channels_repo
        self.generations_repo = generations_repo
        self.generation_access_service = generation_access_service
        self.bot_username = bot_username

    def audience_label(self, filter_key: str) -> str:
        return self.AUDIENCE_FILTERS.get(filter_key, self.AUDIENCE_FILTERS['all'])

    def build_audience_query(self, filter_key: str) -> dict[str, Any]:
        now = utcnow()
        if filter_key == 'subscribed':
            return {'subscription_verified': True}
        if filter_key == 'unsubscribed':
            return {'subscription_verified': {'$ne': True}}
        if filter_key == 'active_7d':
            return {'last_active_at': {'$gte': now - timedelta(days=7)}}
        if filter_key == 'inactive_30d':
            return {'last_active_at': {'$lt': now - timedelta(days=30)}}
        if filter_key == 'referrers':
            return {'referral_count': {'$gt': 0}}
        if filter_key == 'generation_blocked':
            return {'generation_access_blocked': True}
        if filter_key == 'bot_blocked':
            return {'bot_access_blocked': True}
        if filter_key == 'no_credits':
            return {
                'generation_unlimited': {'$ne': True},
                'generation_access_blocked': {'$ne': True},
                'bot_access_blocked': {'$ne': True},
                'free_generation_used': True,
                'referral_credits': {'$lte': 0},
                'bonus_generation_credits': {'$lte': 0},
            }
        return {}

    async def count_audience(self, filter_key: str) -> int:
        return await self.users_repo.count_matching(self.build_audience_query(filter_key))

    async def build_statistics(self) -> dict[str, Any]:
        now = utcnow()
        active_24h_query = {'last_active_at': {'$gte': now - timedelta(days=1)}}
        active_7d_query = {'last_active_at': {'$gte': now - timedelta(days=7)}}
        recent_30d_query = {'created_at': {'$gte': now - timedelta(days=30)}}
        recent_7d_query = {'created_at': {'$gte': now - timedelta(days=7)}}
        recent_24h_query = {'created_at': {'$gte': now - timedelta(days=1)}}

        status_counts = await self.generations_repo.get_status_counts()
        stats = {
            'total_users': await self.users_repo.count_matching(),
            'admins': await self.users_repo.count_matching({'is_admin': True}),
            'subscribed': await self.users_repo.count_matching({'subscription_verified': True}),
            'unsubscribed': await self.users_repo.count_matching({'subscription_verified': {'$ne': True}}),
            'active_24h': await self.users_repo.count_matching(active_24h_query),
            'active_7d': await self.users_repo.count_matching(active_7d_query),
            'new_24h': await self.users_repo.count_matching(recent_24h_query),
            'new_7d': await self.users_repo.count_matching(recent_7d_query),
            'new_30d': await self.users_repo.count_matching(recent_30d_query),
            'referral_total': await self.users_repo.sum_field('referral_count'),
            'manual_bonus_total': await self.users_repo.sum_field('manual_generation_total_added'),
            'generation_unlimited': await self.users_repo.count_matching({'generation_unlimited': True}),
            'generation_blocked': await self.users_repo.count_matching({'generation_access_blocked': True}),
            'bot_blocked': await self.users_repo.count_matching({'bot_access_blocked': True}),
            'total_generated': await self.users_repo.sum_field('generated_count'),
            'total_successful': await self.users_repo.sum_field('successful_generations'),
            'generation_statuses': status_counts,
            'done_24h': await self.generations_repo.count_matching({'status': 'done', 'finished_at': {'$gte': now - timedelta(days=1)}}),
            'done_7d': await self.generations_repo.count_matching({'status': 'done', 'finished_at': {'$gte': now - timedelta(days=7)}}),
            'failed_7d': await self.generations_repo.count_matching({'status': 'failed', 'finished_at': {'$gte': now - timedelta(days=7)}}),
            'recent_failures': await self.generations_repo.list_recent_failures(limit=5),
        }
        return stats

    async def build_ratings(self) -> dict[str, Any]:
        return {
            'top_referrers': await self.users_repo.list_top_referrers(limit=10),
            'top_generators': await self.users_repo.list_top_generators(limit=10),
        }

    async def search_users(self, query: str, limit: int = 10) -> list[dict[str, Any]]:
        return await self.users_repo.search_users(query, limit=limit)

    async def build_user_card(self, telegram_id: int) -> dict[str, Any] | None:
        user = await self.users_repo.get_by_telegram_id(int(telegram_id))
        if not user:
            return None
        available_generations = self.generation_access_service.available_generations(user)
        active_job = await self.generations_repo.get_active_job_for_user(int(telegram_id))
        queue_ahead = 0
        if active_job:
            queue_ahead = await self.generations_repo.count_ahead_in_queue(active_job['_id'])
        return {
            'user': user,
            'available_generations': available_generations,
            'active_job': active_job,
            'queue_ahead': queue_ahead,
        }

    def extract_draft_from_message(self, message: Message) -> dict[str, Any] | None:
        if message.text:
            return {
                'kind': 'text',
                'text': message.text,
            }
        if message.photo:
            return {
                'kind': 'photo',
                'file_id': message.photo[-1].file_id,
                'caption': message.caption or '',
            }
        if message.document:
            document: Document = message.document
            return {
                'kind': 'document',
                'file_id': document.file_id,
                'caption': message.caption or '',
                'filename': document.file_name or 'file',
            }
        if message.video:
            return {
                'kind': 'video',
                'file_id': message.video.file_id,
                'caption': message.caption or '',
            }
        if message.animation:
            return {
                'kind': 'animation',
                'file_id': message.animation.file_id,
                'caption': message.caption or '',
            }
        return None

    def parse_buttons(self, raw_text: str) -> list[dict[str, str]]:
        buttons: list[dict[str, str]] = []
        lines = [line.strip() for line in (raw_text or '').splitlines() if line.strip()]
        for line in lines:
            if '|' not in line:
                raise ValueError('Har bir tugma satri “Matn | manzil” formatida bo‘lishi kerak.')
            text, value = [part.strip() for part in line.split('|', 1)]
            if not text or not value:
                raise ValueError('Tugma matni va manzili bo‘sh bo‘lmasligi kerak.')

            if value.startswith(('http://', 'https://', 'tg://')):
                buttons.append({'text': text, 'type': 'url', 'value': value})
                continue

            if value.startswith('callback:'):
                action = value.removeprefix('callback:').strip().lower()
                if action not in self.PUBLIC_BUTTON_ACTIONS:
                    allowed = ', '.join(sorted(self.PUBLIC_BUTTON_ACTIONS))
                    raise ValueError(f'Callback noto‘g‘ri. Ruxsat etilganlari: {allowed}')
                buttons.append({'text': text, 'type': 'callback', 'value': action})
                continue

            raise ValueError('Manzil URL yoki callback:* ko‘rinishida bo‘lishi kerak.')
        return buttons

    def build_inline_markup(self, buttons: list[dict[str, str]] | None) -> InlineKeyboardMarkup | None:
        if not buttons:
            return None

        rows: list[list[InlineKeyboardButton]] = []
        for button in buttons:
            if button['type'] == 'url':
                rows.append([InlineKeyboardButton(text=button['text'], url=button['value'])])
            else:
                rows.append(
                    [
                        InlineKeyboardButton(
                            text=button['text'],
                            callback_data=PublicPostCallback(action=button['value']).pack(),
                        )
                    ]
                )
        return InlineKeyboardMarkup(inline_keyboard=rows)

    async def send_draft(
        self,
        *,
        bot: Bot,
        chat_id: int,
        draft: dict[str, Any],
        buttons: list[dict[str, str]] | None = None,
    ):
        reply_markup = self.build_inline_markup(buttons)
        kind = draft.get('kind')
        if kind == 'text':
            return await bot.send_message(
                chat_id=chat_id,
                text=draft.get('text', ''),
                parse_mode=ParseMode.HTML,
                reply_markup=reply_markup,
                disable_web_page_preview=False,
            )
        if kind == 'photo':
            return await bot.send_photo(
                chat_id=chat_id,
                photo=draft['file_id'],
                caption=draft.get('caption') or None,
                parse_mode=ParseMode.HTML,
                reply_markup=reply_markup,
            )
        if kind == 'document':
            return await bot.send_document(
                chat_id=chat_id,
                document=draft['file_id'],
                caption=draft.get('caption') or None,
                parse_mode=ParseMode.HTML,
                reply_markup=reply_markup,
            )
        if kind == 'video':
            return await bot.send_video(
                chat_id=chat_id,
                video=draft['file_id'],
                caption=draft.get('caption') or None,
                parse_mode=ParseMode.HTML,
                reply_markup=reply_markup,
            )
        if kind == 'animation':
            return await bot.send_animation(
                chat_id=chat_id,
                animation=draft['file_id'],
                caption=draft.get('caption') or None,
                parse_mode=ParseMode.HTML,
                reply_markup=reply_markup,
            )
        raise ValueError('Qo‘llab-quvvatlanmaydigan xabar turi.')

    async def broadcast(
        self,
        *,
        bot: Bot,
        filter_key: str,
        draft: dict[str, Any],
        buttons: list[dict[str, str]] | None = None,
    ) -> dict[str, int]:
        audience_query = self.build_audience_query(filter_key)
        success_count = 0
        failed_count = 0
        processed = 0

        async for user in self.users_repo.iterate_matching(audience_query, sort=[('created_at', DESCENDING)]):
            processed += 1
            try:
                await self.send_draft(bot=bot, chat_id=int(user['telegram_id']), draft=draft, buttons=buttons)
                success_count += 1
            except Exception:
                failed_count += 1

            if processed % 20 == 0:
                await asyncio.sleep(0.15)

        return {
            'processed': processed,
            'success': success_count,
            'failed': failed_count,
        }

    async def export_users(self, *, filter_key: str, fmt: str = 'csv') -> tuple[str, int]:
        audience_query = self.build_audience_query(filter_key)
        users = await self.users_repo.list_matching(audience_query, limit=100000, sort=[('created_at', DESCENDING)])
        rows = [self._user_export_row(user) for user in users]

        if fmt == 'xlsx':
            return self._export_xlsx(rows, filter_key), len(rows)
        return self._export_csv(rows, filter_key), len(rows)

    def _user_export_row(self, user: dict[str, Any]) -> dict[str, Any]:
        return {
            'telegram_id': user.get('telegram_id', ''),
            'username': user.get('username') or '',
            'full_name': user.get('full_name') or '',
            'subscription_verified': bool(user.get('subscription_verified')),
            'referral_count': int(user.get('referral_count', 0) or 0),
            'referral_credits': int(user.get('referral_credits', 0) or 0),
            'bonus_generation_credits': int(user.get('bonus_generation_credits', 0) or 0),
            'generation_unlimited': bool(user.get('generation_unlimited')),
            'generation_access_blocked': bool(user.get('generation_access_blocked')),
            'bot_access_blocked': bool(user.get('bot_access_blocked')),
            'generated_count': int(user.get('generated_count', 0) or 0),
            'successful_generations': int(user.get('successful_generations', 0) or 0),
            'invited_by': user.get('invited_by') or '',
            'created_at': self._iso(user.get('created_at')),
            'last_active_at': self._iso(user.get('last_active_at')),
        }

    def _export_csv(self, rows: list[dict[str, Any]], filter_key: str) -> str:
        file = tempfile.NamedTemporaryFile(prefix=f'users_{filter_key}_', suffix='.csv', delete=False)
        path = file.name
        file.close()
        fieldnames = list(rows[0].keys()) if rows else [
            'telegram_id', 'username', 'full_name', 'subscription_verified', 'referral_count', 'referral_credits',
            'bonus_generation_credits', 'generation_unlimited', 'generation_access_blocked', 'bot_access_blocked',
            'generated_count', 'successful_generations', 'invited_by', 'created_at', 'last_active_at',
        ]
        with open(path, 'w', encoding='utf-8-sig', newline='') as csv_file:
            writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(rows)
        return path

    def _export_xlsx(self, rows: list[dict[str, Any]], filter_key: str) -> str:
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = 'users'
        headers = list(rows[0].keys()) if rows else [
            'telegram_id', 'username', 'full_name', 'subscription_verified', 'referral_count', 'referral_credits',
            'bonus_generation_credits', 'generation_unlimited', 'generation_access_blocked', 'bot_access_blocked',
            'generated_count', 'successful_generations', 'invited_by', 'created_at', 'last_active_at',
        ]
        sheet.append(headers)
        for cell in sheet[1]:
            cell.font = Font(bold=True)
        for row in rows:
            sheet.append([row.get(header, '') for header in headers])
        for column_cells in sheet.columns:
            max_length = max(len(str(cell.value or '')) for cell in column_cells)
            sheet.column_dimensions[column_cells[0].column_letter].width = min(max_length + 2, 36)

        file = tempfile.NamedTemporaryFile(prefix=f'users_{filter_key}_', suffix='.xlsx', delete=False)
        path = file.name
        file.close()
        workbook.save(path)
        return path

    @staticmethod
    def _iso(value: Any) -> str:
        if not value:
            return ''
        try:
            return value.isoformat()
        except Exception:
            return str(value)

    def cleanup_file(self, path: str) -> None:
        try:
            Path(path).unlink(missing_ok=True)
        except Exception:
            pass

    async def list_mandatory_channels(self) -> list[dict[str, Any]]:
        return await self.channels_repo.list_all()

    async def get_mandatory_channel(self, chat_id: int) -> dict[str, Any] | None:
        return await self.channels_repo.get_by_chat_id(chat_id)

    async def set_channel_active(self, chat_id: int, value: bool) -> dict[str, Any] | None:
        return await self.channels_repo.set_active(chat_id, value)

    async def delete_channel(self, chat_id: int) -> bool:
        return await self.channels_repo.delete_by_chat_id(chat_id)

    async def save_mandatory_channel(self, channel_payload: dict[str, Any]) -> dict[str, Any]:
        return await self.channels_repo.upsert_channel(
            chat_id=int(channel_payload['chat_id']),
            title=str(channel_payload['title']),
            username=channel_payload.get('username'),
            invite_link=channel_payload.get('invite_link'),
            is_active=bool(channel_payload.get('is_active', True)),
        )

    async def resolve_channel_reference(self, *, bot: Bot, raw_reference: str) -> dict[str, Any]:
        reference = self._normalize_channel_reference(raw_reference)
        if not reference:
            raise ValueError('Kanal manzili noto‘g‘ri. @username, t.me havola yoki chat_id yuboring.')

        try:
            chat = await bot.get_chat(reference)
        except Exception as exc:
            raise ValueError('Kanal topilmadi yoki bot uni ko‘ra olmayapti.') from exc

        if getattr(chat, 'type', None) not in {'channel', 'supergroup'}:
            raise ValueError('Faqat kanal yoki supergroup turidagi chat qo‘llab-quvvatlanadi.')

        me = await bot.get_me()
        try:
            await bot.get_chat_member(chat_id=chat.id, user_id=me.id)
        except Exception as exc:
            raise ValueError(
                'Bot ushbu kanalga qo‘shilmagan yoki obunani tekshirish uchun yetarli huquqqa ega emas.'
            ) from exc

        username = getattr(chat, 'username', None)
        username = username.lstrip('@') if username else None
        invite_link = f'https://t.me/{username}' if username else None
        return {
            'chat_id': int(chat.id),
            'title': getattr(chat, 'title', None) or username or str(chat.id),
            'username': username,
            'invite_link': invite_link,
            'is_active': True,
        }

    @staticmethod
    def normalize_invite_link(value: str) -> str:
        link = str(value or '').strip()
        if not link:
            raise ValueError('Invite link bo‘sh bo‘lmasligi kerak.')
        if link.startswith('http://'):
            link = 'https://' + link.removeprefix('http://')
        if link.startswith('t.me/'):
            link = 'https://' + link
        if not re.match(r'^https://t\.me/[^\s]+$', link):
            raise ValueError('Invite link https://t.me/... ko‘rinishida bo‘lishi kerak.')
        return link

    @staticmethod
    def _normalize_channel_reference(value: str) -> str | int | None:
        raw = str(value or '').strip()
        if not raw:
            return None
        if raw.startswith('https://t.me/'):
            raw = raw.removeprefix('https://t.me/')
        elif raw.startswith('http://t.me/'):
            raw = raw.removeprefix('http://t.me/')
        elif raw.startswith('t.me/'):
            raw = raw.removeprefix('t.me/')
        raw = raw.strip().rstrip('/')
        if raw.startswith('@'):
            return raw
        if raw.lstrip('-').isdigit():
            return int(raw)
        if re.fullmatch(r'[A-Za-z0-9_]{4,}', raw):
            return f'@{raw}'
        return None
