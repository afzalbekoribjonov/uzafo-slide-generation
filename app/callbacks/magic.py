from aiogram.filters.callback_data import CallbackData


class MagicMenuCallback(CallbackData, prefix='mgu'):
    action: str
    value: str = '0'


class MagicAdminCallback(CallbackData, prefix='mga'):
    action: str
    value: str = '0'


class MagicCardCallback(CallbackData, prefix='mgc'):
    action: str
    card_id: str


class MagicTopupCallback(CallbackData, prefix='mgt'):
    action: str
    topup_id: str
