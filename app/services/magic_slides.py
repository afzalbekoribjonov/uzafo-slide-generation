from __future__ import annotations

import json
from html import escape
from typing import Any

from aiogram import Bot
from aiogram.types import Message

from app.repositories.magic_accounts import MagicAccountsRepository
from app.repositories.magic_cards import MagicCardsRepository
from app.repositories.magic_orders import MagicOrdersRepository
from app.repositories.magic_settings import MagicSettingsRepository
from app.repositories.magic_topups import MagicTopupsRepository


class MagicSlideService:
    TOPUP_AMOUNTS = (5000, 10000, 15000, 25000, 50000)

    def __init__(
        self,
        *,
        settings_repo: MagicSettingsRepository,
        cards_repo: MagicCardsRepository,
        accounts_repo: MagicAccountsRepository,
        topups_repo: MagicTopupsRepository,
        orders_repo: MagicOrdersRepository,
        admin_ids: set[int],
        webapp_url: str | None,
    ) -> None:
        self.settings_repo = settings_repo
        self.cards_repo = cards_repo
        self.accounts_repo = accounts_repo
        self.topups_repo = topups_repo
        self.orders_repo = orders_repo
        self.admin_ids = set(admin_ids or set())
        self.webapp_url = str(webapp_url or '').strip() or None

    async def get_user_context(self, telegram_id: int) -> dict[str, Any]:
        settings = await self.settings_repo.get_settings()
        account = await self.accounts_repo.ensure_account(telegram_id)
        cards = await self.cards_repo.list_active()
        price = int(settings.get('price_per_presentation', MagicSettingsRepository.DEFAULT_PRICE_PER_PRESENTATION) or 0)
        balance = int(account.get('balance_uzs', 0) or 0)
        available_presentations = balance // price if price > 0 else 0
        return {
            'settings': settings,
            'account': account,
            'cards': cards,
            'price_uzs': price,
            'balance_uzs': balance,
            'available_presentations': available_presentations,
            'can_afford': price > 0 and balance >= price,
            'maintenance_enabled': bool(settings.get('maintenance_enabled')),
            'webapp_url': self.webapp_url,
            'webapp_configured': bool(self.webapp_url),
        }

    async def get_settings_context(self) -> dict[str, Any]:
        settings = await self.settings_repo.get_settings()
        cards = await self.cards_repo.list_all()
        pending_count = await self.topups_repo.count_pending()
        return {
            'settings': settings,
            'cards': cards,
            'pending_count': pending_count,
            'webapp_configured': bool(self.webapp_url),
        }

    async def list_cards(self) -> list[dict[str, Any]]:
        return await self.cards_repo.list_all()

    async def list_pending_topups(self, limit: int = 20) -> list[dict[str, Any]]:
        return await self.topups_repo.list_pending(limit=limit)

    async def get_card(self, card_id: str) -> dict[str, Any] | None:
        return await self.cards_repo.get_by_id(card_id)

    async def get_topup(self, topup_id: str) -> dict[str, Any] | None:
        return await self.topups_repo.get_by_id(topup_id)

    async def set_price(self, amount_uzs: int) -> dict[str, Any]:
        if int(amount_uzs) <= 0:
            raise ValueError('Narx 0 dan katta bo‘lishi kerak.')
        return await self.settings_repo.set_price(int(amount_uzs))

    async def toggle_maintenance(self) -> dict[str, Any]:
        settings = await self.settings_repo.get_settings()
        return await self.settings_repo.set_maintenance(not bool(settings.get('maintenance_enabled')))

    async def create_card(self, raw_value: str) -> dict[str, Any]:
        card_number, card_holder = self.parse_card_details(raw_value)
        return await self.cards_repo.create_card(card_number=card_number, card_holder=card_holder, is_active=True)

    async def toggle_card(self, card_id: str) -> dict[str, Any] | None:
        card = await self.cards_repo.get_by_id(card_id)
        if not card:
            return None
        return await self.cards_repo.set_active(card_id, not bool(card.get('is_active')))

    async def delete_card(self, card_id: str) -> bool:
        return await self.cards_repo.delete_by_id(card_id)

    async def create_topup_request(self, *, user: dict[str, Any], amount_uzs: int, receipt: dict[str, Any]) -> dict[str, Any]:
        amount_uzs = int(amount_uzs)
        if amount_uzs not in self.TOPUP_AMOUNTS:
            raise ValueError('Noto‘g‘ri to‘lov summasi tanlandi.')

        cards = await self.cards_repo.list_active()
        if not cards:
            raise ValueError('Hozircha faol qabul kartalari mavjud emas.')

        return await self.topups_repo.create_pending(
            telegram_id=int(user['telegram_id']),
            full_name=str(user.get('full_name') or user['telegram_id']),
            username=user.get('username'),
            amount_uzs=amount_uzs,
            receipt_type=str(receipt['kind']),
            receipt_file_id=str(receipt['file_id']),
            receipt_file_unique_id=receipt.get('file_unique_id'),
            receipt_file_name=receipt.get('file_name'),
            receipt_mime_type=receipt.get('mime_type'),
            receipt_caption=receipt.get('caption'),
            cards_snapshot=self.cards_snapshot(cards),
        )

    async def notify_admins_about_topup(self, bot: Bot, topup: dict[str, Any], reply_markup) -> int:
        sent_count = 0
        for admin_id in self.admin_ids:
            try:
                review_message = await self.send_topup_for_review(bot, chat_id=admin_id, topup=topup, reply_markup=reply_markup)
            except Exception:
                continue

            if review_message:
                await self.topups_repo.add_admin_notification(
                    str(topup['_id']),
                    chat_id=admin_id,
                    message_id=review_message.message_id,
                )
                sent_count += 1
        return sent_count

    async def resend_topup_to_admin(self, *, bot: Bot, chat_id: int, topup_id: str, reply_markup):
        topup = await self.topups_repo.get_by_id(topup_id)
        if not topup:
            return None

        review_message = await self.send_topup_for_review(bot, chat_id=chat_id, topup=topup, reply_markup=reply_markup)
        if review_message:
            await self.topups_repo.add_admin_notification(
                topup_id,
                chat_id=chat_id,
                message_id=review_message.message_id,
            )
        return topup

    async def send_topup_for_review(self, bot: Bot, *, chat_id: int, topup: dict[str, Any], reply_markup):
        caption = self.build_admin_topup_caption(topup)
        if topup.get('receipt_type') == 'photo':
            return await bot.send_photo(
                chat_id=chat_id,
                photo=topup['receipt_file_id'],
                caption=caption,
                reply_markup=reply_markup,
            )
        return await bot.send_document(
            chat_id=chat_id,
            document=topup['receipt_file_id'],
            caption=caption,
            reply_markup=reply_markup,
        )

    async def clear_admin_review_keyboards(self, bot: Bot, topup: dict[str, Any]) -> None:
        for notification in topup.get('admin_notifications') or []:
            try:
                await bot.edit_message_reply_markup(
                    chat_id=int(notification['chat_id']),
                    message_id=int(notification['message_id']),
                    reply_markup=None,
                )
            except Exception:
                continue

    async def approve_topup(self, *, topup_id: str, admin_id: int, admin_name: str) -> dict[str, Any] | None:
        topup = await self.topups_repo.mark_approved(topup_id, admin_id=admin_id, admin_name=admin_name)
        if not topup:
            return None

        account = await self.accounts_repo.add_balance(int(topup['telegram_id']), int(topup['amount_uzs']))
        topup['account_balance_uzs'] = int(account.get('balance_uzs', 0) or 0)
        return topup

    async def reject_topup(self, *, topup_id: str, admin_id: int, admin_name: str) -> dict[str, Any] | None:
        return await self.topups_repo.mark_rejected(topup_id, admin_id=admin_id, admin_name=admin_name)

    async def create_order_draft(self, *, user: dict[str, Any], raw_payload: str) -> dict[str, Any]:
        payload = self.parse_webapp_payload(raw_payload)
        settings = await self.settings_repo.get_settings()
        return await self.orders_repo.create_job(
            telegram_id=int(user['telegram_id']),
            full_name=str(user.get('full_name') or user['telegram_id']),
            username=user.get('username'),
            payload=payload,
            template_id=str(payload.get('template_id') or ''),
            template_name=str(payload.get('template_name') or payload.get('template_id') or 'Magic Slide'),
            category=str(payload.get('category') or 'general'),
            output_slide_target=int(payload.get('output_slide_target', 0) or 0),
            price_uzs_snapshot=int(
                settings.get('price_per_presentation', MagicSettingsRepository.DEFAULT_PRICE_PER_PRESENTATION) or 0
            ),
        )

    async def create_order_job(
        self,
        *,
        user: dict[str, Any],
        raw_payload: str,
    ) -> tuple[dict[str, Any] | None, int, dict[str, Any] | None]:
        existing = await self.orders_repo.get_active_order_for_user(int(user['telegram_id']))
        if existing:
            ahead_count = await self.orders_repo.count_ahead_in_queue(existing['_id'])
            return None, ahead_count, existing

        context = await self.get_user_context(int(user['telegram_id']))
        if context['maintenance_enabled']:
            raise ValueError('Hozircha yangi buyurtma qabul qilinmayapti. Birozdan keyin qayta urinib ko‘ring.')
        if not context['can_afford']:
            raise ValueError("Hisobingizdagi mablag‘ hozircha yetarli emas. Avval hisobingizni to‘ldirib, keyin qayta urinib ko‘ring.")

        order = await self.create_order_draft(user=user, raw_payload=raw_payload)
        ahead_count = await self.orders_repo.count_ahead_in_queue(order['_id'])
        return order, ahead_count, None

    async def set_order_status_message(self, order_id: str, *, chat_id: int, message_id: int) -> None:
        await self.orders_repo.set_status_message(order_id, chat_id=chat_id, message_id=message_id)

    @staticmethod
    def parse_receipt_message(message: Message) -> dict[str, Any]:
        if message.photo:
            photo = message.photo[-1]
            return {
                'kind': 'photo',
                'file_id': photo.file_id,
                'file_unique_id': photo.file_unique_id,
                'file_name': None,
                'mime_type': 'image/jpeg',
                'caption': message.caption,
            }

        if message.document:
            return {
                'kind': 'document',
                'file_id': message.document.file_id,
                'file_unique_id': message.document.file_unique_id,
                'file_name': message.document.file_name,
                'mime_type': message.document.mime_type,
                'caption': message.caption,
            }

        raise ValueError('Chekni rasm yoki PDF/document ko‘rinishida yuboring.')

    @staticmethod
    def parse_webapp_payload(raw_payload: str) -> dict[str, Any]:
        try:
            payload = json.loads(raw_payload)
        except json.JSONDecodeError as exc:
            raise ValueError('WebApp dan kelgan JSON o‘qilmadi.') from exc

        if not isinstance(payload, dict):
            raise ValueError('WebApp payload noto‘g‘ri formatda keldi.')
        if payload.get('flow') != 'magic_slide':
            raise ValueError('WebApp payload Magic Slayd oqimiga tegishli emas.')
        if not str(payload.get('template_id') or '').strip():
            raise ValueError('Payload ichida template_id topilmadi.')
        if not isinstance(payload.get('variables'), dict):
            raise ValueError('Payload ichidagi variables bo‘limi topilmadi.')
        if 'extra_context' in payload and not isinstance(payload.get('extra_context'), dict):
            raise ValueError('Payload ichidagi extra_context noto‘g‘ri formatda keldi.')
        return payload

    @staticmethod
    def parse_card_details(raw_value: str) -> tuple[str, str]:
        parts = [part.strip() for part in str(raw_value or '').split('|', 1)]
        if len(parts) != 2 or not parts[0] or not parts[1]:
            raise ValueError('Kartani <code>8600 1234 5678 9012 | CARD HOLDER</code> ko‘rinishida yuboring.')
        return MagicSlideService.normalize_card_number(parts[0]), MagicSlideService.normalize_card_holder(parts[1])

    @staticmethod
    def normalize_card_number(value: str) -> str:
        digits = ''.join(ch for ch in str(value or '') if ch.isdigit())
        if len(digits) < 12 or len(digits) > 19:
            raise ValueError('Karta raqami 12 dan 19 gacha raqamdan iborat bo‘lishi kerak.')
        return digits

    @staticmethod
    def format_card_number(value: str) -> str:
        digits = ''.join(ch for ch in str(value or '') if ch.isdigit())
        if not digits:
            return ''
        return ' '.join(digits[index:index + 4] for index in range(0, len(digits), 4))

    @staticmethod
    def normalize_card_holder(value: str) -> str:
        holder = ' '.join(str(value or '').strip().split())
        if len(holder) < 3:
            raise ValueError('Karta egasi ismini to‘liqroq yozing.')
        return holder

    @staticmethod
    def format_money(amount: int | float | None) -> str:
        return f"{int(amount or 0):,}".replace(',', ' ')

    @classmethod
    def mask_card_number(cls, card_number: str) -> str:
        digits = ''.join(ch for ch in str(card_number or '') if ch.isdigit())
        if len(digits) <= 4:
            return digits
        masked = '*' * max(0, len(digits) - 4) + digits[-4:]
        return ' '.join(masked[index:index + 4] for index in range(0, len(masked), 4))

    @classmethod
    def cards_snapshot(cls, cards: list[dict[str, Any]]) -> list[dict[str, Any]]:
        snapshot: list[dict[str, Any]] = []
        for card in cards:
            snapshot.append(
                {
                    'card_id': str(card['_id']),
                    'card_holder': str(card.get('card_holder') or ''),
                    'masked_number': cls.mask_card_number(str(card.get('card_number') or '')),
                }
            )
        return snapshot

    @classmethod
    def payment_cards_snapshot(cls, cards: list[dict[str, Any]]) -> list[dict[str, Any]]:
        snapshot: list[dict[str, Any]] = []
        for card in cards:
            raw_number = str(card.get('card_number') or '')
            normalized_number = cls.normalize_card_number(raw_number) if raw_number else ''
            snapshot.append(
                {
                    'card_id': str(card['_id']),
                    'card_holder': str(card.get('card_holder') or ''),
                    'full_number': cls.format_card_number(normalized_number),
                    'copy_number': normalized_number,
                }
            )
        return snapshot

    @classmethod
    def build_admin_topup_caption(cls, topup: dict[str, Any]) -> str:
        username = topup.get('username')
        username_text = f"@{escape(str(username).lstrip('@'))}" if username else '—'
        cards = topup.get('cards_snapshot') or []
        card_lines = []
        for index, card in enumerate(cards, start=1):
            card_lines.append(
                f"{index}. <code>{escape(str(card.get('masked_number', '—')))}</code> — <b>{escape(str(card.get('card_holder', '—')))}</b>"
            )
        cards_text = '\n'.join(card_lines) if card_lines else '—'
        return (
            "<b>💳 Magic Slayd to‘lov cheki</b>\n\n"
            f"• Foydalanuvchi: <b>{escape(str(topup.get('full_name', 'Noma’lum')))}</b>\n"
            f"• Telegram ID: <code>{topup.get('telegram_id')}</code>\n"
            f"• Username: <b>{username_text}</b>\n"
            f"• To‘lov summasi: <b>{cls.format_money(topup.get('amount_uzs'))} so‘m</b>\n"
            f"• Yuborilgan vaqt: <b>{escape(str(topup.get('created_at')))}</b>\n\n"
            "<u>Ko‘rsatilgan kartalar</u>\n"
            f"{cards_text}"
        )
