from __future__ import annotations

from aiogram.types import (
    CopyTextButton,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    KeyboardButton,
    ReplyKeyboardMarkup,
    WebAppInfo,
)
from aiogram.utils.keyboard import InlineKeyboardBuilder

from app.callbacks.admin import AdminMenuCallback
from app.callbacks.magic import MagicAdminCallback, MagicCardCallback, MagicMenuCallback, MagicTopupCallback
from app.callbacks.menu import MenuCallback


def magic_home_keyboard() -> InlineKeyboardMarkup:
    builder = InlineKeyboardBuilder()
    builder.button(text='💼 Hisobim', callback_data=MagicMenuCallback(action='account'))
    builder.button(text='🚀 Yaratishni boshlash', callback_data=MagicMenuCallback(action='start'))
    builder.button(text='◀️ Bosh menyuga qaytish', callback_data=MenuCallback(action='main'))
    builder.adjust(2, 1)
    return builder.as_markup()


def magic_account_keyboard() -> InlineKeyboardMarkup:
    builder = InlineKeyboardBuilder()
    builder.button(text='◀️ Ortga', callback_data=MagicMenuCallback(action='home'))
    builder.button(text='💳 Hisobni to‘ldirish', callback_data=MagicMenuCallback(action='topup'))
    builder.button(text='◀️ Bosh menyuga qaytish', callback_data=MenuCallback(action='main'))
    builder.adjust(2, 1)
    return builder.as_markup()


def magic_topup_amount_keyboard() -> InlineKeyboardMarkup:
    builder = InlineKeyboardBuilder()
    for amount in (5000, 10000, 15000, 25000, 50000):
        builder.button(
            text=f"{amount:,}".replace(',', '.'),
            callback_data=MagicMenuCallback(action='amount', value=str(amount)),
        )
    builder.button(text='◀️ Ortga qaytish', callback_data=MagicMenuCallback(action='account'))
    builder.button(text='◀️ Bosh menyuga qaytish', callback_data=MenuCallback(action='main'))
    builder.adjust(3, 2, 1, 1)
    return builder.as_markup()


def magic_receipt_wait_keyboard(cards: list[dict] | None = None) -> InlineKeyboardMarkup:
    builder = InlineKeyboardBuilder()
    for index, card in enumerate(cards or [], start=1):
        copy_number = str(card.get('copy_number') or '').strip()
        if not copy_number:
            continue
        builder.row(
            InlineKeyboardButton(
                text=f'📋 {index}-kartani nusxalash',
                copy_text=CopyTextButton(text=copy_number),
            )
        )
    builder.row(InlineKeyboardButton(text='◀️ Ortga qaytish', callback_data=MagicMenuCallback(action='topup').pack()))
    builder.row(InlineKeyboardButton(text='◀️ Bosh menyuga qaytish', callback_data=MenuCallback(action='main').pack()))
    return builder.as_markup()


def magic_start_keyboard(webapp_url: str) -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        keyboard=[
            [
                KeyboardButton(
                    text='Slayd yaratishni boshlash',
                    web_app=WebAppInfo(url=webapp_url),
                )
            ]
        ],
        resize_keyboard=True,
        one_time_keyboard=True,
        input_field_placeholder='Pastdagi tugma orqali formani oching',
    )


def magic_start_blocked_keyboard() -> InlineKeyboardMarkup:
    builder = InlineKeyboardBuilder()
    builder.button(text='💼 Hisobim', callback_data=MagicMenuCallback(action='account'))
    builder.button(text='💳 Hisobni to‘ldirish', callback_data=MagicMenuCallback(action='topup'))
    builder.button(text='◀️ Bosh menyuga qaytish', callback_data=MenuCallback(action='main'))
    builder.adjust(2, 1)
    return builder.as_markup()


def admin_magic_settings_keyboard(*, maintenance_enabled: bool) -> InlineKeyboardMarkup:
    builder = InlineKeyboardBuilder()
    builder.button(text='💰 Narxni yangilash', callback_data=MagicAdminCallback(action='price'))
    builder.button(
        text='🛠 Maintenance: ON' if maintenance_enabled else '✅ Maintenance: OFF',
        callback_data=MagicAdminCallback(action='maintenance_toggle'),
    )
    builder.button(text='💳 Kartalar', callback_data=MagicAdminCallback(action='cards'))
    builder.button(text='🧾 Kutilayotgan to‘lovlar', callback_data=MagicAdminCallback(action='pending'))
    builder.button(text='◀️ Admin menyu', callback_data=AdminMenuCallback(action='main'))
    builder.adjust(2, 2, 1)
    return builder.as_markup()


def admin_magic_cards_keyboard(cards: list[dict]) -> InlineKeyboardMarkup:
    builder = InlineKeyboardBuilder()
    for card in cards[:20]:
        status_icon = '🟢' if card.get('is_active') else '⚪️'
        holder = str(card.get('card_holder') or 'Karta')
        masked = str(card.get('masked_number') or '—')
        builder.button(
            text=f'{status_icon} {holder[:18]} • {masked[-4:]}',
            callback_data=MagicCardCallback(action='open', card_id=str(card['_id'])),
        )
    builder.button(text='➕ Karta qo‘shish', callback_data=MagicAdminCallback(action='cards_add'))
    builder.button(text='◀️ Sozlamalarga qaytish', callback_data=MagicAdminCallback(action='settings'))
    builder.adjust(1)
    return builder.as_markup()


def admin_magic_card_keyboard(card: dict) -> InlineKeyboardMarkup:
    card_id = str(card['_id'])
    builder = InlineKeyboardBuilder()
    builder.button(
        text='⏸ Faolsiz qilish' if card.get('is_active') else '▶️ Faollashtirish',
        callback_data=MagicCardCallback(action='toggle', card_id=card_id),
    )
    builder.button(text='🗑 O‘chirish', callback_data=MagicCardCallback(action='delete', card_id=card_id))
    builder.button(text='🔄 Yangilash', callback_data=MagicCardCallback(action='open', card_id=card_id))
    builder.button(text='◀️ Kartalar ro‘yxati', callback_data=MagicAdminCallback(action='cards'))
    builder.button(text='◀️ Sozlamalar', callback_data=MagicAdminCallback(action='settings'))
    builder.adjust(2, 1, 1, 1)
    return builder.as_markup()


def admin_magic_pending_keyboard(topups: list[dict]) -> InlineKeyboardMarkup:
    builder = InlineKeyboardBuilder()
    for topup in topups[:20]:
        name = str(topup.get('full_name') or topup.get('telegram_id'))
        amount = f"{int(topup.get('amount_uzs', 0) or 0):,}".replace(',', '.')
        builder.button(
            text=f'🧾 {amount} • {name[:18]}',
            callback_data=MagicTopupCallback(action='open', topup_id=str(topup['_id'])),
        )
    builder.button(text='🔄 Yangilash', callback_data=MagicAdminCallback(action='pending'))
    builder.button(text='◀️ Sozlamalar', callback_data=MagicAdminCallback(action='settings'))
    builder.adjust(1)
    return builder.as_markup()


def admin_magic_topup_review_keyboard(topup_id: str) -> InlineKeyboardMarkup:
    builder = InlineKeyboardBuilder()
    builder.button(text='✅ Tasdiqlayman', callback_data=MagicTopupCallback(action='approve', topup_id=topup_id))
    builder.button(text='❌ To‘lov soxta', callback_data=MagicTopupCallback(action='reject', topup_id=topup_id))
    builder.adjust(2)
    return builder.as_markup()
