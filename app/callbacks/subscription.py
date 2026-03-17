from aiogram.filters.callback_data import CallbackData


class SubscriptionCallback(CallbackData, prefix='sub'):
    action: str