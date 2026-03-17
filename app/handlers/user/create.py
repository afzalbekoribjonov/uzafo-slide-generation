from __future__ import annotations

from aiogram import F, Router
from aiogram.exceptions import TelegramBadRequest
from aiogram.fsm.context import FSMContext
from aiogram.types import CallbackQuery, Message

from app.callbacks.menu import CreateFlowCallback, MenuCallback
from app.keyboards.user import (
    create_confirm_keyboard,
    create_credit_missing_keyboard,
    create_language_keyboard,
    create_slide_count_keyboard,
    main_menu_keyboard,
)
from app.repositories.users import UsersRepository
from app.services.generations import GenerationAccessService
from app.services.generation_queue import GenerationQueueService
from app.states.create import CreatePresentationStates
from app.texts.user import (
    create_already_queued_text,
    create_confirmation_text,
    create_credit_missing_text,
    create_generation_failed_text,
    create_generation_blocked_text,
    create_language_prompt_text,
    create_presenter_prompt_text,
    create_queued_text,
    create_slide_count_prompt_text,
    create_topic_prompt_text,
    create_validation_error_text,
    main_menu_text,
)

router = Router(name='user-create')


@router.callback_query(MenuCallback.filter(F.action == 'create'))
async def create_entry_handler(
    callback: CallbackQuery,
    state: FSMContext,
    users_repo: UsersRepository,
    generation_access_service: GenerationAccessService,
    generation_queue_service: GenerationQueueService,
) -> None:
    user = await users_repo.get_by_telegram_id(callback.from_user.id)
    if not user:
        await callback.answer('Foydalanuvchi topilmadi.', show_alert=True)
        return

    if user.get('generation_access_blocked'):
        await state.clear()
        await callback.message.edit_text(
            text=create_generation_blocked_text(),
            reply_markup=create_credit_missing_keyboard(back_only=True),
        )
        await callback.answer('Generation imkoniyati cheklangan.', show_alert=True)
        return

    if not generation_access_service.has_available_generation(user):
        await state.clear()
        await callback.message.edit_text(
            text=create_credit_missing_text(),
            reply_markup=create_credit_missing_keyboard(),
        )
        await callback.answer()
        return

    existing_job, ahead_count = await generation_queue_service.describe_existing_job(callback.from_user.id)
    if existing_job:
        await state.clear()
        await callback.message.edit_text(
            text=create_already_queued_text(ahead_count),
            reply_markup=main_menu_keyboard(is_admin=bool(user.get('is_admin'))),
        )
        await callback.answer('Sizda faol so‘rov mavjud.', show_alert=True)
        return

    await state.clear()
    await state.set_state(CreatePresentationStates.waiting_topic)
    await callback.message.edit_text(
        text=create_topic_prompt_text(),
        reply_markup=create_credit_missing_keyboard(back_only=True),
    )
    await callback.answer()


@router.message(CreatePresentationStates.waiting_topic)
async def create_topic_handler(message: Message, state: FSMContext) -> None:
    topic = (message.text or '').strip()
    if len(topic) < 3:
        await message.answer(create_validation_error_text('Mavzu kamida 3 ta belgidan iborat bo‘lishi kerak.'))
        return

    await state.update_data(topic=topic)
    await state.set_state(CreatePresentationStates.waiting_presenter_name)
    await message.answer(create_presenter_prompt_text())


@router.message(CreatePresentationStates.waiting_presenter_name)
async def create_presenter_name_handler(message: Message, state: FSMContext) -> None:
    presenter_name = (message.text or '').strip()
    if len(presenter_name) < 2:
        await message.answer(create_validation_error_text('Tayyorlagan ism kamida 2 ta belgidan iborat bo‘lishi kerak.'))
        return

    await state.update_data(presenter_name=presenter_name)
    await state.set_state(CreatePresentationStates.waiting_slide_count)
    await message.answer(
        text=create_slide_count_prompt_text(),
        reply_markup=create_slide_count_keyboard(),
    )


@router.callback_query(CreateFlowCallback.filter(F.action == 'slides'))
async def create_slide_count_handler(callback: CallbackQuery, callback_data: CreateFlowCallback, state: FSMContext) -> None:
    current_state = await state.get_state()
    if current_state != CreatePresentationStates.waiting_slide_count.state:
        await callback.answer()
        return

    slide_count = int(callback_data.value)
    if slide_count < 6 or slide_count > 15:
        await callback.answer('Faqat 6 dan 15 gacha tanlang.', show_alert=True)
        return

    await state.update_data(slide_count=slide_count)
    await state.set_state(CreatePresentationStates.waiting_language)

    try:
        await callback.message.edit_text(
            text=create_language_prompt_text(),
            reply_markup=create_language_keyboard(),
        )
    except TelegramBadRequest:
        await callback.message.answer(
            text=create_language_prompt_text(),
            reply_markup=create_language_keyboard(),
        )
    await callback.answer()


@router.callback_query(CreateFlowCallback.filter(F.action == 'language'))
async def create_language_handler(callback: CallbackQuery, callback_data: CreateFlowCallback, state: FSMContext) -> None:
    current_state = await state.get_state()
    if current_state != CreatePresentationStates.waiting_language.state:
        await callback.answer()
        return

    language_map = {
        'uz': 'O‘zbek',
        'ru': 'Русский',
        'en': 'English',
    }
    language_code = callback_data.value
    language_name = language_map.get(language_code)
    if not language_name:
        await callback.answer('Noto‘g‘ri til tanlandi.', show_alert=True)
        return

    await state.update_data(language_code=language_code, language_name=language_name)
    data = await state.get_data()
    await state.set_state(CreatePresentationStates.waiting_confirmation)

    try:
        await callback.message.edit_text(
            text=create_confirmation_text(data),
            reply_markup=create_confirm_keyboard(),
        )
    except TelegramBadRequest:
        await callback.message.answer(
            text=create_confirmation_text(data),
            reply_markup=create_confirm_keyboard(),
        )
    await callback.answer()


@router.callback_query(CreateFlowCallback.filter(F.action == 'confirm'))
async def create_confirm_handler(
    callback: CallbackQuery,
    callback_data: CreateFlowCallback,
    state: FSMContext,
    users_repo: UsersRepository,
    generation_access_service: GenerationAccessService,
    generation_queue_service: GenerationQueueService,
) -> None:
    action = callback_data.value
    user = await users_repo.get_by_telegram_id(callback.from_user.id)

    if action == 'cancel':
        await state.clear()
        await callback.message.edit_text(
            text=main_menu_text(user['full_name']),
            reply_markup=main_menu_keyboard(is_admin=bool(user.get('is_admin'))),
        )
        await callback.answer('Jarayon bekor qilindi.')
        return

    current_state = await state.get_state()
    if current_state != CreatePresentationStates.waiting_confirmation.state:
        await callback.answer()
        return

    existing_job, ahead_count = await generation_queue_service.describe_existing_job(callback.from_user.id)
    if existing_job:
        await state.clear()
        await callback.message.edit_text(
            text=create_already_queued_text(ahead_count),
            reply_markup=main_menu_keyboard(is_admin=bool(user.get('is_admin'))),
        )
        await callback.answer('Sizda faol so‘rov mavjud.', show_alert=True)
        return

    if user.get('generation_access_blocked'):
        await state.clear()
        await callback.message.edit_text(
            text=create_generation_blocked_text(),
            reply_markup=create_credit_missing_keyboard(back_only=True),
        )
        await callback.answer('Generation imkoniyati cheklangan.', show_alert=True)
        return

    if not user or not generation_access_service.has_available_generation(user):
        await state.clear()
        await callback.message.edit_text(
            text=create_credit_missing_text(),
            reply_markup=create_credit_missing_keyboard(),
        )
        await callback.answer('Yetarli limit topilmadi.', show_alert=True)
        return

    consumed_from = await generation_access_service.consume_generation(users_repo, callback.from_user.id)
    if not consumed_from:
        await state.clear()
        await callback.message.edit_text(
            text=create_credit_missing_text(),
            reply_markup=create_credit_missing_keyboard(),
        )
        await callback.answer('Limit sarflanmadi.', show_alert=True)
        return

    data = await state.get_data()
    try:
        job, ahead_count, active_job = await generation_queue_service.create_job(
            telegram_id=callback.from_user.id,
            full_name=user['full_name'],
            username=callback.from_user.username,
            payload=data,
            consumed_from=consumed_from,
            status_chat_id=callback.message.chat.id,
            status_message_id=callback.message.message_id,
        )
    except Exception:
        await generation_access_service.restore_consumed_generation(users_repo, callback.from_user.id, consumed_from)
        await state.clear()
        await callback.message.edit_text(
            text=create_generation_failed_text(),
            reply_markup=main_menu_keyboard(is_admin=bool(user.get('is_admin'))),
        )
        await callback.answer('So‘rov saqlanmadi.', show_alert=True)
        return

    if active_job:
        await generation_access_service.restore_consumed_generation(users_repo, callback.from_user.id, consumed_from)
        await state.clear()
        await callback.message.edit_text(
            text=create_already_queued_text(ahead_count),
            reply_markup=main_menu_keyboard(is_admin=bool(user.get('is_admin'))),
        )
        await callback.answer('Sizda faol so‘rov mavjud.', show_alert=True)
        return

    await state.clear()

    try:
        await callback.message.edit_text(
            text=create_queued_text(data, ahead_count),
            reply_markup=main_menu_keyboard(is_admin=bool(user.get('is_admin'))),
        )
    except TelegramBadRequest:
        sent = await callback.message.answer(
            text=create_queued_text(data, ahead_count),
            reply_markup=main_menu_keyboard(is_admin=bool(user.get('is_admin'))),
        )
        if job:
            await generation_queue_service.generations_repo.set_status_message(
                job['_id'],
                chat_id=sent.chat.id,
                message_id=sent.message_id,
            )
    await callback.answer('So‘rov navbatga qo‘shildi.')


@router.callback_query(CreateFlowCallback.filter(F.action == 'cancel'))
async def create_cancel_handler(callback: CallbackQuery, state: FSMContext, users_repo: UsersRepository) -> None:
    await state.clear()
    user = await users_repo.get_by_telegram_id(callback.from_user.id)
    await callback.message.edit_text(
        text=main_menu_text(user['full_name']),
        reply_markup=main_menu_keyboard(is_admin=bool(user.get('is_admin'))),
    )
    await callback.answer('Jarayon bekor qilindi.')
