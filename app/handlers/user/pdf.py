from __future__ import annotations

import logging
from pathlib import Path
from aiogram import Bot, Router, F
from aiogram.types import CallbackQuery, FSInputFile

from app.callbacks.pdf import PdfConvertCallback
from app.repositories.generations import GenerationsRepository
from app.services.pdf_converter import PdfConverterService

logger = logging.getLogger(__name__)
router = Router(name='user-pdf')

@router.callback_query(PdfConvertCallback.filter())
async def pdf_convert_handler(
    callback: CallbackQuery,
    callback_data: PdfConvertCallback,
    generations_repo: GenerationsRepository,
    bot: Bot,
) -> None:
    action = callback_data.action
    generation_id = callback_data.generation_id
    pptx_msg_id = callback_data.pptx_msg_id
    
    if action == 'no':
        await callback.answer('Bekor qilindi.')
        try:
            # Delete messages
            await bot.delete_message(chat_id=callback.message.chat.id, message_id=pptx_msg_id)
            await callback.message.delete()
            
            # Cleanup file
            job = await generations_repo.get_by_id(generation_id)
            if job and job.get('result_file_path'):
                path = Path(job['result_file_path'])
                if path.exists():
                    path.unlink()
                    logger.info(f"File {path} deleted after user clicked 'No'")
        except Exception as e:
            logger.warning(f"Error during 'No' callback cleanup: {e}")
        return

    if action == 'yes':
        await callback.answer('PDF tayyorlanmoqda...', show_alert=False)
        await generations_repo.set_pdf_processing(generation_id, True)
        
        job = await generations_repo.get_by_id(generation_id)
        
        if not job or not job.get('result_file_path'):
            await callback.message.answer("Kechirasiz, fayl topilmadi yoki muddati o'tgan.")
            await callback.message.delete()
            return
            
        pptx_path = job['result_file_path']
        pdf_path = await PdfConverterService.convert_to_pdf(pptx_path)
        
        if pdf_path:
            try:
                await bot.send_document(
                    chat_id=callback.message.chat.id,
                    document=FSInputFile(pdf_path),
                    caption="Tayyorlangan PDF shakli."
                )
                # Cleanup PDF after sending
                Path(pdf_path).unlink()
            except Exception as e:
                logger.error(f"Error sending PDF or cleaning up: {e}")
        else:
            await callback.message.answer("PDFga o'girishda xatolik yuz berdi. Iltimos keyinroq urinib ko'ring.")
            
        # Cleanup PPTX and messages anyway after 'yes' processing
        try:
            if Path(pptx_path).exists():
                Path(pptx_path).unlink()
            await bot.delete_message(chat_id=callback.message.chat.id, message_id=pptx_msg_id)
            await callback.message.delete()
        except Exception as e:
            logger.warning(f"Error during 'Yes' callback final cleanup: {e}")
