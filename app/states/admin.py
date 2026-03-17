from aiogram.fsm.state import State, StatesGroup


class AdminUserSearchStates(StatesGroup):
    waiting_query = State()
    waiting_credit_amount = State()


class AdminBroadcastStates(StatesGroup):
    waiting_content = State()
    waiting_buttons = State()
