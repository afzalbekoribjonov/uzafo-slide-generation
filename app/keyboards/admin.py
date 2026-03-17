from __future__ import annotations

from aiogram.types import InlineKeyboardMarkup
from aiogram.utils.keyboard import InlineKeyboardBuilder

from app.callbacks.admin import AdminBroadcastCallback, AdminMenuCallback, AdminUserCallback
from app.services.admin import AdminService
from app.callbacks.menu import MenuCallback



def admin_main_keyboard() -> InlineKeyboardMarkup:
    builder = InlineKeyboardBuilder()
    builder.button(text='📈 Statistika', callback_data=AdminMenuCallback(action='stats'))
    builder.button(text='🏆 Reyting', callback_data=AdminMenuCallback(action='rating'))
    builder.button(text='🔎 Foydalanuvchi qidirish', callback_data=AdminMenuCallback(action='users'))
    builder.button(text='📣 Ommaviy xabar', callback_data=AdminMenuCallback(action='broadcast'))
    builder.button(text='📤 Eksport', callback_data=AdminMenuCallback(action='exports'))
    builder.button(text='◀️ User menyusi', callback_data=MenuCallback(action='main'))
    builder.adjust(2, 1, 1, 1, 1)
    return builder.as_markup()



def admin_secondary_keyboard() -> InlineKeyboardMarkup:
    builder = InlineKeyboardBuilder()
    builder.button(text='◀️ Admin menyuga qaytish', callback_data=AdminMenuCallback(action='main'))
    return builder.as_markup()



def admin_search_results_keyboard(results: list[dict]) -> InlineKeyboardMarkup:
    builder = InlineKeyboardBuilder()
    for user in results:
        name = user.get('full_name') or str(user.get('telegram_id'))
        builder.button(
            text=f"👤 {name[:24]}",
            callback_data=AdminUserCallback(action='open', user_id=int(user['telegram_id'])),
        )
    builder.button(text='🔄 Yangi qidiruv', callback_data=AdminMenuCallback(action='users'))
    builder.button(text='◀️ Admin menyu', callback_data=AdminMenuCallback(action='main'))
    builder.adjust(1)
    return builder.as_markup()



def admin_user_card_keyboard(user: dict) -> InlineKeyboardMarkup:
    telegram_id = int(user['telegram_id'])
    builder = InlineKeyboardBuilder()
    builder.button(text='➕ Kredit qo‘shish', callback_data=AdminUserCallback(action='credit_add', user_id=telegram_id))
    builder.button(text='➖ Kredit ayirish', callback_data=AdminUserCallback(action='credit_remove', user_id=telegram_id))
    builder.button(
        text='♾ Cheksizni o‘chirish' if user.get('generation_unlimited') else '♾ Cheksiz qilish',
        callback_data=AdminUserCallback(action='toggle_unlimited', user_id=telegram_id),
    )
    builder.button(
        text='✅ Generation ochish' if user.get('generation_access_blocked') else '🚫 Generation bloklash',
        callback_data=AdminUserCallback(action='toggle_generation_block', user_id=telegram_id),
    )
    builder.button(
        text='♻️ Bot kirishini ochish' if user.get('bot_access_blocked') else '⛔ Bot kirishini bloklash',
        callback_data=AdminUserCallback(action='toggle_bot_block', user_id=telegram_id),
    )
    builder.button(text='🔄 Yangilash', callback_data=AdminUserCallback(action='open', user_id=telegram_id))
    builder.button(text='🔎 Qidiruvga qaytish', callback_data=AdminMenuCallback(action='users'))
    builder.button(text='◀️ Admin menyu', callback_data=AdminMenuCallback(action='main'))
    builder.adjust(2, 2, 1, 1, 1)
    return builder.as_markup()



def admin_broadcast_audience_keyboard() -> InlineKeyboardMarkup:
    builder = InlineKeyboardBuilder()
    for filter_key, label in AdminService.AUDIENCE_FILTERS.items():
        builder.button(text=label, callback_data=AdminBroadcastCallback(action='audience', value=filter_key))
    builder.button(text='◀️ Admin menyu', callback_data=AdminMenuCallback(action='main'))
    builder.adjust(1)
    return builder.as_markup()



def admin_broadcast_skip_buttons_keyboard() -> InlineKeyboardMarkup:
    builder = InlineKeyboardBuilder()
    builder.button(text='⏭ O‘tkazib yuborish', callback_data=AdminBroadcastCallback(action='buttons_skip'))
    builder.button(text='◀️ Admin menyu', callback_data=AdminMenuCallback(action='main'))
    builder.adjust(1, 1)
    return builder.as_markup()



def admin_broadcast_preview_keyboard() -> InlineKeyboardMarkup:
    builder = InlineKeyboardBuilder()
    builder.button(text='✏️ Matnni almashtirish', callback_data=AdminBroadcastCallback(action='edit_content'))
    builder.button(text='🔘 Tugmalarni tahrirlash', callback_data=AdminBroadcastCallback(action='edit_buttons'))
    builder.button(text='🧪 Test yuborish', callback_data=AdminBroadcastCallback(action='test'))
    builder.button(text='🚀 Barchaga yuborish', callback_data=AdminBroadcastCallback(action='send'))
    builder.button(text='🗑 Bekor qilish', callback_data=AdminBroadcastCallback(action='cancel'))
    builder.button(text='◀️ Admin menyu', callback_data=AdminMenuCallback(action='main'))
    builder.adjust(2, 2, 1, 1)
    return builder.as_markup()



def admin_export_keyboard() -> InlineKeyboardMarkup:
    builder = InlineKeyboardBuilder()
    for filter_key, label in AdminService.AUDIENCE_FILTERS.items():
        builder.button(text=label, callback_data=AdminMenuCallback(action='export_audience', value=filter_key))
    builder.button(text='◀️ Admin menyu', callback_data=AdminMenuCallback(action='main'))
    builder.adjust(1)
    return builder.as_markup()



def admin_export_format_keyboard(filter_key: str) -> InlineKeyboardMarkup:
    builder = InlineKeyboardBuilder()
    builder.button(text='📄 CSV', callback_data=AdminMenuCallback(action='export_format', value=f'{filter_key}__csv'))
    builder.button(text='📘 XLSX', callback_data=AdminMenuCallback(action='export_format', value=f'{filter_key}__xlsx'))
    builder.button(text='◀️ Auditoriyalar', callback_data=AdminMenuCallback(action='exports'))
    builder.adjust(2, 1)
    return builder.as_markup()
