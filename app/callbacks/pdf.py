from aiogram.filters.callback_data import CallbackData

class PdfConvertCallback(CallbackData, prefix='pdf'):
    action: str  # 'yes', 'no'
    generation_id: str
    pptx_msg_id: int
