from aiogram.fsm.state import State, StatesGroup


class AdminUserSearchStates(StatesGroup):
    waiting_query = State()
    waiting_credit_amount = State()


class AdminBroadcastStates(StatesGroup):
    waiting_content = State()
    waiting_buttons = State()


class AdminChannelStates(StatesGroup):
    waiting_channel_reference = State()
    waiting_channel_invite_link = State()
