from aiogram.filters.callback_data import CallbackData


class AdminMenuCallback(CallbackData, prefix='adm'):
    action: str
    value: str = '0'


class AdminUserCallback(CallbackData, prefix='admu'):
    action: str
    user_id: int
    value: str = '0'


class AdminBroadcastCallback(CallbackData, prefix='admb'):
    action: str
    value: str = '0'


class AdminChannelCallback(CallbackData, prefix='admc'):
    action: str
    chat_id: int
    value: str = '0'


class PublicPostCallback(CallbackData, prefix='pub'):
    action: str
