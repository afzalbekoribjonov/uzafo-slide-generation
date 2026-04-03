from __future__ import annotations

from aiogram import F, Router
from aiogram.fsm.context import FSMContext
from aiogram.types import CallbackQuery, Message, ReplyKeyboardRemove

from app.callbacks.magic import MagicMenuCallback
from app.callbacks.menu import MenuCallback
from app.keyboards.magic import (
    MAGIC_START_CANCEL_TEXT,
    admin_magic_topup_review_keyboard,
    magic_account_keyboard,
    magic_home_keyboard,
    magic_receipt_wait_keyboard,
    magic_start_blocked_keyboard,
    magic_start_keyboard,
    magic_topup_amount_keyboard,
)
from app.repositories.users import UsersRepository
from app.services.magic_slides import MagicSlideService
from app.states.magic import MagicTopUpStates
from app.states.magic import MagicOrderStates
from app.texts.magic import (
    magic_account_text,
    magic_hook_text,
    magic_maintenance_text,
    magic_order_existing_text,
    magic_order_queued_text,
    magic_receipt_prompt_text,
    magic_receipt_received_text,
    magic_start_cancelled_text,
    magic_start_insufficient_text,
    magic_start_prompt_text,
    magic_start_ready_text,
    magic_topup_approved_text,
    magic_topup_rejected_text,
    magic_topup_text,
    magic_topup_unavailable_text,
    magic_webapp_not_ready_text,
)

router = Router(name='user-magic')


async def _delete_magic_prompt_message(
    *,
    bot,
    state: FSMContext,
) -> bool:
    data = await state.get_data()
    prompt_chat_id = data.get('magic_webapp_prompt_chat_id')
    prompt_message_id = data.get('magic_webapp_prompt_message_id')
    had_prompt = bool(prompt_chat_id and prompt_message_id)
    if had_prompt:
        try:
            await bot.delete_message(chat_id=int(prompt_chat_id), message_id=int(prompt_message_id))
        except Exception:
            pass
    await state.update_data(magic_webapp_prompt_chat_id=None, magic_webapp_prompt_message_id=None)
    return had_prompt


async def _dismiss_magic_start_flow(
    *,
    bot,
    state: FSMContext,
    chat_id: int,
    notify: bool = False,
) -> bool:
    had_prompt = await _delete_magic_prompt_message(bot=bot, state=state)
    current_state = await state.get_state()
    if current_state == MagicOrderStates.waiting_webapp.state:
        await state.clear()
    if had_prompt and notify:
        await bot.send_message(
            chat_id=chat_id,
            text=magic_start_cancelled_text(),
            reply_markup=ReplyKeyboardRemove(),
        )
    return had_prompt


@router.callback_query(MenuCallback.filter(F.action == 'magic'))
async def magic_menu_entry_handler(
    callback: CallbackQuery,
    state: FSMContext,
    magic_slide_service: MagicSlideService,
) -> None:
    await _dismiss_magic_start_flow(
        bot=callback.bot,
        state=state,
        chat_id=callback.from_user.id,
        notify=True,
    )
    await state.clear()
    context = await magic_slide_service.get_user_context(callback.from_user.id)
    await callback.message.edit_text(
        text=magic_hook_text(context),
        reply_markup=magic_home_keyboard(),
    )
    await callback.answer()


@router.callback_query(MagicMenuCallback.filter(F.action == 'home'))
async def magic_home_handler(
    callback: CallbackQuery,
    state: FSMContext,
    magic_slide_service: MagicSlideService,
) -> None:
    await _dismiss_magic_start_flow(
        bot=callback.bot,
        state=state,
        chat_id=callback.from_user.id,
        notify=True,
    )
    await state.clear()
    context = await magic_slide_service.get_user_context(callback.from_user.id)
    await callback.message.edit_text(
        text=magic_hook_text(context),
        reply_markup=magic_home_keyboard(),
    )
    await callback.answer()


@router.callback_query(MagicMenuCallback.filter(F.action == 'account'))
async def magic_account_handler(
    callback: CallbackQuery,
    state: FSMContext,
    magic_slide_service: MagicSlideService,
) -> None:
    await _dismiss_magic_start_flow(
        bot=callback.bot,
        state=state,
        chat_id=callback.from_user.id,
        notify=True,
    )
    context = await magic_slide_service.get_user_context(callback.from_user.id)
    await callback.message.edit_text(
        text=magic_account_text(context),
        reply_markup=magic_account_keyboard(),
    )
    await callback.answer()


@router.callback_query(MagicMenuCallback.filter(F.action == 'topup'))
async def magic_topup_handler(
    callback: CallbackQuery,
    state: FSMContext,
    magic_slide_service: MagicSlideService,
) -> None:
    await _dismiss_magic_start_flow(
        bot=callback.bot,
        state=state,
        chat_id=callback.from_user.id,
        notify=True,
    )
    await state.clear()
    context = await magic_slide_service.get_user_context(callback.from_user.id)
    if not context['cards']:
        await callback.message.edit_text(
            text=magic_topup_unavailable_text(),
            reply_markup=magic_account_keyboard(),
        )
        await callback.answer()
        return

    payment_cards = magic_slide_service.payment_cards_snapshot(context['cards'])
    await callback.message.edit_text(
        text=magic_topup_text(payment_cards),
        reply_markup=magic_topup_amount_keyboard(),
    )
    await callback.answer()


@router.callback_query(MagicMenuCallback.filter(F.action == 'amount'))
async def magic_amount_selected_handler(
    callback: CallbackQuery,
    callback_data: MagicMenuCallback,
    state: FSMContext,
    magic_slide_service: MagicSlideService,
) -> None:
    amount_uzs = int(callback_data.value or 0)
    context = await magic_slide_service.get_user_context(callback.from_user.id)
    if amount_uzs not in magic_slide_service.TOPUP_AMOUNTS:
        await callback.answer('Noto‘g‘ri summa tanlandi.', show_alert=True)
        return

    if not context['cards']:
        await state.clear()
        await callback.message.edit_text(
            text=magic_topup_unavailable_text(),
            reply_markup=magic_account_keyboard(),
        )
        await callback.answer()
        return

    payment_cards = magic_slide_service.payment_cards_snapshot(context['cards'])
    await state.set_state(MagicTopUpStates.waiting_receipt)
    await state.update_data(magic_topup_amount=amount_uzs)
    await callback.message.edit_text(
        text=magic_receipt_prompt_text(amount_uzs, payment_cards),
        reply_markup=magic_receipt_wait_keyboard(payment_cards),
    )
    await callback.answer()


@router.message(MagicTopUpStates.waiting_receipt, F.photo | F.document)
async def magic_receipt_handler(
    message: Message,
    state: FSMContext,
    users_repo: UsersRepository,
    magic_slide_service: MagicSlideService,
) -> None:
    data = await state.get_data()
    amount_uzs = int(data.get('magic_topup_amount', 0) or 0)
    if not amount_uzs:
        await state.clear()
        await message.answer('Tanlangan to‘lov summasi topilmadi. Qaytadan urinib ko‘ring.')
        return

    try:
        receipt = magic_slide_service.parse_receipt_message(message)
    except ValueError as exc:
        await message.answer(str(exc))
        return

    user = await users_repo.get_by_telegram_id(message.from_user.id)
    if not user:
        await state.clear()
        await message.answer('Foydalanuvchi topilmadi. /start orqali qayta kirib ko‘ring.')
        return

    topup = await magic_slide_service.create_topup_request(user=user, amount_uzs=amount_uzs, receipt=receipt)
    await state.clear()

    await magic_slide_service.notify_admins_about_topup(
        message.bot,
        topup,
        reply_markup=admin_magic_topup_review_keyboard(str(topup['_id'])),
    )

    await message.answer(
        text=magic_receipt_received_text(amount_uzs),
        reply_markup=magic_account_keyboard(),
    )


@router.message(MagicTopUpStates.waiting_receipt)
async def magic_receipt_invalid_handler(message: Message) -> None:
    await message.answer('Chekni rasm yoki PDF/document ko‘rinishida yuboring.')


@router.callback_query(MagicMenuCallback.filter(F.action == 'start'))
async def magic_start_handler(
    callback: CallbackQuery,
    state: FSMContext,
    magic_slide_service: MagicSlideService,
) -> None:
    await _dismiss_magic_start_flow(
        bot=callback.bot,
        state=state,
        chat_id=callback.from_user.id,
        notify=False,
    )
    context = await magic_slide_service.get_user_context(callback.from_user.id)

    if context['maintenance_enabled']:
        await callback.message.edit_text(
            text=magic_maintenance_text(),
            reply_markup=magic_home_keyboard(),
        )
        await callback.answer()
        return

    if not context['can_afford']:
        await callback.message.edit_text(
            text=magic_start_insufficient_text(context),
            reply_markup=magic_start_blocked_keyboard(),
        )
        await callback.answer()
        return

    if not context['webapp_configured']:
        await callback.message.edit_text(
            text=magic_webapp_not_ready_text(),
            reply_markup=magic_home_keyboard(),
        )
        await callback.answer()
        return

    prompt_message = await callback.message.answer(
        text=f"{magic_start_ready_text(context)}\n\n{magic_start_prompt_text()}",
        reply_markup=magic_start_keyboard(context['webapp_url']),
    )
    await state.set_state(MagicOrderStates.waiting_webapp)
    await state.update_data(
        magic_webapp_prompt_chat_id=prompt_message.chat.id,
        magic_webapp_prompt_message_id=prompt_message.message_id,
    )
    await callback.answer()


@router.message(MagicOrderStates.waiting_webapp, F.text == MAGIC_START_CANCEL_TEXT)
async def magic_start_cancel_handler(
    message: Message,
    state: FSMContext,
    magic_slide_service: MagicSlideService,
) -> None:
    await _dismiss_magic_start_flow(
        bot=message.bot,
        state=state,
        chat_id=message.chat.id,
        notify=False,
    )
    context = await magic_slide_service.get_user_context(message.from_user.id)
    await message.answer(
        magic_start_cancelled_text(),
        reply_markup=ReplyKeyboardRemove(),
    )
    await message.answer(
        magic_hook_text(context),
        reply_markup=magic_home_keyboard(),
    )


@router.message(MagicOrderStates.waiting_webapp, F.text)
async def magic_start_waiting_handler(message: Message) -> None:
    await message.answer(
        "Pastdagi tugma orqali yaratishni boshlang yoki bekor qilish tugmasini bosing."
    )


@router.message(F.web_app_data)
async def magic_webapp_data_handler(
    message: Message,
    state: FSMContext,
    users_repo: UsersRepository,
    magic_slide_service: MagicSlideService,
) -> None:
    await _delete_magic_prompt_message(bot=message.bot, state=state)

    user = await users_repo.get_by_telegram_id(message.from_user.id)
    if not user:
        await state.clear()
        await message.answer('Foydalanuvchi topilmadi. /start orqali qayta kirib ko‘ring.', reply_markup=ReplyKeyboardRemove())
        return

    context = await magic_slide_service.get_user_context(message.from_user.id)
    if context['maintenance_enabled']:
        await state.clear()
        await message.answer(magic_maintenance_text(), reply_markup=ReplyKeyboardRemove())
        return
    if not context['can_afford']:
        await state.clear()
        await message.answer(magic_start_insufficient_text(context), reply_markup=ReplyKeyboardRemove())
        return

    try:
        order, ahead_count, existing = await magic_slide_service.create_order_job(
            user=user,
            raw_payload=message.web_app_data.data,
        )
    except ValueError as exc:
        await state.clear()
        await message.answer(str(exc), reply_markup=ReplyKeyboardRemove())
        return

    if existing:
        await state.clear()
        await message.answer(
            magic_order_existing_text(
                template_name=str(existing.get('template_name') or ''),
                ahead_count=ahead_count,
            ),
            reply_markup=ReplyKeyboardRemove(),
        )
        return

    await state.clear()
    status_message = await message.answer(
        magic_order_queued_text(
            template_name=str(order.get('template_name') or 'Magic Slide'),
            ahead_count=ahead_count,
        ),
        reply_markup=ReplyKeyboardRemove(),
    )
    await magic_slide_service.set_order_status_message(
        str(order['_id']),
        chat_id=status_message.chat.id,
        message_id=status_message.message_id,
    )


async def notify_magic_topup_approved(bot, telegram_id: int, amount_uzs: int, balance_uzs: int) -> None:
    await bot.send_message(
        telegram_id,
        magic_topup_approved_text(amount_uzs, balance_uzs),
        reply_markup=magic_account_keyboard(),
    )


async def notify_magic_topup_rejected(bot, telegram_id: int, amount_uzs: int) -> None:
    await bot.send_message(
        telegram_id,
        magic_topup_rejected_text(amount_uzs),
        reply_markup=magic_account_keyboard(),
    )
