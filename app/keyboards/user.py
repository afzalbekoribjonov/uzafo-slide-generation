from aiogram.types import InlineKeyboardMarkup
from aiogram.utils.keyboard import InlineKeyboardBuilder

from app.callbacks.admin import AdminMenuCallback
from app.callbacks.menu import CreateFlowCallback, MenuCallback, StatusCallback
from app.callbacks.subscription import SubscriptionCallback



def main_menu_keyboard(*, is_admin: bool = False) -> InlineKeyboardMarkup:
    builder = InlineKeyboardBuilder()
    builder.button(text='🎞 Slayd yaratish', callback_data=MenuCallback(action='create'))
    builder.button(text='✨ Magic slayd', callback_data=MenuCallback(action='magic'))
    builder.button(text='📊 Mening holatim', callback_data=MenuCallback(action='status'))
    builder.button(text='👥 Taklif qilish', callback_data=MenuCallback(action='invite'))
    builder.button(text='❓ Yordam', callback_data=MenuCallback(action='help'))
    builder.button(text='☎️ Aloqa', callback_data=MenuCallback(action='contact'))
    if is_admin:
        builder.button(text='🛡 Admin panel', callback_data=AdminMenuCallback(action='main'))
        builder.adjust(2, 2, 2, 1)
    else:
        builder.adjust(2, 2, 2)
    return builder.as_markup()



def subscription_keyboard(channels: list[dict]) -> InlineKeyboardMarkup:
    builder = InlineKeyboardBuilder()

    for index, channel in enumerate(channels, start=1):
        title = channel.get('title') or channel.get('username') or f'Kanal {index}'
        url = channel.get('invite_link')

        if not url and channel.get('username'):
            username = str(channel['username']).lstrip('@')
            url = f'https://t.me/{username}'

        if url:
            builder.button(text=f'📢 Obuna bo‘lish — {title}', url=url)

    builder.button(text='✅ Obunani tekshirish', callback_data=SubscriptionCallback(action='check'))
    builder.adjust(1)
    return builder.as_markup()



def status_keyboard() -> InlineKeyboardMarkup:
    builder = InlineKeyboardBuilder()
    builder.button(text='👥 Takliflar ro‘yxati', callback_data=StatusCallback(action='referrals'))
    builder.button(text='◀️ Ortga qaytish', callback_data=MenuCallback(action='main'))
    builder.adjust(1, 1)
    return builder.as_markup()



def invite_keyboard() -> InlineKeyboardMarkup:
    builder = InlineKeyboardBuilder()
    builder.button(text='👥 Takliflar ro‘yxati', callback_data=StatusCallback(action='referrals'))
    builder.button(text='◀️ Ortga qaytish', callback_data=MenuCallback(action='main'))
    builder.adjust(1, 1)
    return builder.as_markup()



def referrals_keyboard() -> InlineKeyboardMarkup:
    builder = InlineKeyboardBuilder()
    builder.button(text='◀️ Ortga qaytish', callback_data=MenuCallback(action='status'))
    return builder.as_markup()



def help_keyboard() -> InlineKeyboardMarkup:
    builder = InlineKeyboardBuilder()
    builder.button(text='◀️ Ortga qaytish', callback_data=MenuCallback(action='main'))
    return builder.as_markup()



def contact_keyboard() -> InlineKeyboardMarkup:
    builder = InlineKeyboardBuilder()
    builder.button(text='◀️ Ortga qaytish', callback_data=MenuCallback(action='main'))
    return builder.as_markup()



def create_credit_missing_keyboard(*, back_only: bool = False) -> InlineKeyboardMarkup:
    builder = InlineKeyboardBuilder()
    if not back_only:
        builder.button(text='👥 Taklif qilish', callback_data=MenuCallback(action='invite'))
    builder.button(text='◀️ Ortga qaytish', callback_data=MenuCallback(action='main'))
    builder.adjust(1)
    return builder.as_markup()



def create_slide_count_keyboard() -> InlineKeyboardMarkup:
    builder = InlineKeyboardBuilder()
    for count in range(6, 16):
        builder.button(text=str(count), callback_data=CreateFlowCallback(action='slides', value=str(count)))
    builder.button(text='❌ Bekor qilish', callback_data=CreateFlowCallback(action='cancel', value='cancel'))
    builder.adjust(5, 5, 1)
    return builder.as_markup()



def create_language_keyboard() -> InlineKeyboardMarkup:
    builder = InlineKeyboardBuilder()
    builder.button(text='🇺🇿 O‘zbek', callback_data=CreateFlowCallback(action='language', value='uz'))
    builder.button(text='🇷🇺 Русский', callback_data=CreateFlowCallback(action='language', value='ru'))
    builder.button(text='🇺🇸 English', callback_data=CreateFlowCallback(action='language', value='en'))
    builder.button(text='❌ Bekor qilish', callback_data=CreateFlowCallback(action='cancel', value='cancel'))
    builder.adjust(1, 1, 1, 1)
    return builder.as_markup()



def create_confirm_keyboard() -> InlineKeyboardMarkup:
    builder = InlineKeyboardBuilder()
    builder.button(text='✅ Tasdiqlash', callback_data=CreateFlowCallback(action='confirm', value='confirm'))
    builder.button(text='❌ Bekor qilish', callback_data=CreateFlowCallback(action='confirm', value='cancel'))
    builder.adjust(1, 1)
    return builder.as_markup()
