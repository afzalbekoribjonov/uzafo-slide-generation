from aiogram.fsm.state import State, StatesGroup


class CreatePresentationStates(StatesGroup):
    waiting_topic = State()
    waiting_presenter_name = State()
    waiting_slide_count = State()
    waiting_language = State()
    waiting_confirmation = State()
