from aiogram.fsm.state import State, StatesGroup


class MagicTopUpStates(StatesGroup):
    waiting_receipt = State()


class AdminMagicStates(StatesGroup):
    waiting_price = State()
    waiting_card_details = State()
