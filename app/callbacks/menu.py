from aiogram.filters.callback_data import CallbackData


class MenuCallback(CallbackData, prefix='menu'):
    action: str


class StatusCallback(CallbackData, prefix='status'):
    action: str


class InviteCallback(CallbackData, prefix='invite'):
    action: str


class HelpCallback(CallbackData, prefix='help'):
    action: str

class CreateFlowCallback(CallbackData, prefix='create'):
    action: str
    value: str
