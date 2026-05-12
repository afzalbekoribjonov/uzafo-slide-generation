from __future__ import annotations

from aiogram import F, Router
from aiogram.exceptions import TelegramBadRequest
from aiogram.fsm.context import FSMContext
from aiogram.types import CallbackQuery, FSInputFile, Message

from app.callbacks.menu import CreateFlowCallback, MenuCallback
from app.keyboards.user import (
    create_confirm_keyboard,
    create_credit_missing_keyboard,
    create_language_keyboard,
    create_pdf_choice_keyboard,
    create_slide_count_keyboard,
    create_template_keyboard,
    main_menu_keyboard,
)
from app.repositories.users import UsersRepository
from app.services.generations import GenerationAccessService
from app.services.generation_queue import GenerationQueueService
from app.services.content_validator import ContentValidator
from app.services.presentation_templates import PresentationTemplateRegistry
from app.states.create import CreatePresentationStates
from app.texts.user import (
    create_already_queued_text,
    create_confirmation_text,
    create_credit_missing_text,
    create_generation_failed_text,
    create_generation_blocked_text,
    create_language_prompt_text,
    create_pdf_choice_prompt_text,
    create_presenter_prompt_text,
    create_queued_text,
    create_slide_count_prompt_text,
    create_template_preview_caption,
    create_template_prompt_text,
    create_topic_prompt_text,
    create_validation_error_text,
    main_menu_text,
)

router = Router(name='user-create')
template_registry = PresentationTemplateRegistry()


async def _safe_edit_text(message: Message, *, text: str, reply_markup=None) -> None:
    try:
        await message.edit_text(text=text, reply_markup=reply_markup)
    except TelegramBadRequest as exc:
        if 'message is not modified' in str(exc).lower():
            return
        raise


async def _show_main_menu_from_callback(
    callback: CallbackQuery,
    *,
    state: FSMContext,
    users_repo: UsersRepository,
    answer_text: str,
) -> None:
    await state.clear()
    user = await users_repo.get_by_telegram_id(callback.from_user.id)
    if not user:
        await callback.answer('Profil topilmadi. /start ni bosib qayta urinib ko‘ring.', show_alert=True)
        return
    text = main_menu_text(user['full_name'])
    markup = main_menu_keyboard(is_admin=bool(user.get('is_admin')))

    try:
        if callback.message.text:
            await callback.message.edit_text(text=text, reply_markup=markup)
        elif callback.message.caption is not None:
            await callback.message.edit_reply_markup(reply_markup=None)
            await callback.message.answer(text=text, reply_markup=markup)
        else:
            await callback.message.answer(text=text, reply_markup=markup)
    except TelegramBadRequest:
        await callback.message.answer(text=text, reply_markup=markup)

    await callback.answer(answer_text)


async def _send_template_selection(message: Message) -> None:
    templates = template_registry.list_templates()
    reply_markup = create_template_keyboard(templates)
    preview_path = template_registry.preview_collage_path
    if preview_path.exists():
        await message.answer_photo(
            photo=FSInputFile(str(preview_path)),
            caption=create_template_preview_caption(),
            reply_markup=reply_markup,
        )
        return

    await message.answer(
        text=create_template_prompt_text(),
        reply_markup=reply_markup,
    )


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
        await callback.answer('Profil topilmadi. /start ni bosib qayta urinib ko‘ring.', show_alert=True)
        return

    if user.get('generation_access_blocked'):
        await state.clear()
        await _safe_edit_text(
            callback.message,
            text=create_generation_blocked_text(),
            reply_markup=create_credit_missing_keyboard(back_only=True),
        )
        await callback.answer('Slayd yaratish hozircha yopiq.', show_alert=True)
        return

    if not generation_access_service.has_available_generation(user):
        await state.clear()
        await _safe_edit_text(
            callback.message,
            text=create_credit_missing_text(),
            reply_markup=create_credit_missing_keyboard(),
        )
        await callback.answer()
        return

    existing_job, ahead_count = await generation_queue_service.describe_existing_job(callback.from_user.id)
    if existing_job:
        await state.clear()
        await _safe_edit_text(
            callback.message,
            text=create_already_queued_text(ahead_count),
            reply_markup=main_menu_keyboard(is_admin=bool(user.get('is_admin'))),
        )
        await callback.answer('Sizda bitta taqdimot tayyorlanmoqda.', show_alert=True)
        return

    await state.clear()
    await state.set_state(CreatePresentationStates.waiting_topic)
    await _safe_edit_text(
        callback.message,
        text=create_topic_prompt_text(),
        reply_markup=create_credit_missing_keyboard(back_only=True),
    )
    await callback.answer()


@router.message(CreatePresentationStates.waiting_topic)
async def create_topic_handler(message: Message, state: FSMContext) -> None:
    topic = (message.text or '').strip()
    is_valid, error_message = ContentValidator.validate_topic(topic)
    if not is_valid:
        await message.answer(create_validation_error_text(error_message))
        return

    await state.update_data(topic=topic)
    await state.set_state(CreatePresentationStates.waiting_presenter_name)
    await message.answer(create_presenter_prompt_text())


@router.message(CreatePresentationStates.waiting_presenter_name)
async def create_presenter_name_handler(message: Message, state: FSMContext) -> None:
    presenter_name = (message.text or '').strip()
    is_valid, error_message = ContentValidator.validate_presenter_name(presenter_name)
    if not is_valid:
        await message.answer(create_validation_error_text(error_message))
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
    if slide_count < 6 or slide_count > 12:
        await callback.answer('6 dan 12 gacha slayd tanlang.', show_alert=True)
        return

    await state.update_data(slide_count=slide_count)
    await state.set_state(CreatePresentationStates.waiting_template)

    try:
        await callback.message.edit_text(text=create_template_prompt_text())
    except TelegramBadRequest:
        await callback.message.answer(text=create_template_prompt_text())

    await _send_template_selection(callback.message)
    await callback.answer()


@router.callback_query(CreateFlowCallback.filter(F.action == 'template'))
async def create_template_handler(callback: CallbackQuery, callback_data: CreateFlowCallback, state: FSMContext) -> None:
    current_state = await state.get_state()
    if current_state != CreatePresentationStates.waiting_template.state:
        await callback.answer()
        return

    template = template_registry.get(callback_data.value)
    await state.update_data(
        template_id=template['id'],
        template_name=template.get('name') or template['id'],
    )
    await state.set_state(CreatePresentationStates.waiting_language)

    # Delete the preview photo message to keep chat clean
    try:
        await callback.message.delete()
    except Exception:
        pass

    await callback.message.answer(
        text=create_language_prompt_text(),
        reply_markup=create_language_keyboard(),
    )
    await callback.answer(f"{template.get('name') or template['id']} dizayni tanlandi.")


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
        await callback.answer('Tilni pastdagi tugmalardan tanlang.', show_alert=True)
        return

    await state.update_data(language_code=language_code, language_name=language_name)
    await state.set_state(CreatePresentationStates.waiting_pdf_choice)

    try:
        await _safe_edit_text(
            callback.message,
            text=create_pdf_choice_prompt_text(),
            reply_markup=create_pdf_choice_keyboard(),
        )
    except TelegramBadRequest:
        await callback.message.answer(
            text=create_pdf_choice_prompt_text(),
            reply_markup=create_pdf_choice_keyboard(),
        )
    await callback.answer()


@router.callback_query(CreateFlowCallback.filter(F.action == 'pdf_choice'))
async def create_pdf_choice_handler(callback: CallbackQuery, callback_data: CreateFlowCallback, state: FSMContext) -> None:
    current_state = await state.get_state()
    if current_state != CreatePresentationStates.waiting_pdf_choice.state:
        await callback.answer()
        return

    wants_pdf = callback_data.value == 'yes'
    await state.update_data(wants_pdf=wants_pdf)
    data = await state.get_data()
    await state.set_state(CreatePresentationStates.waiting_confirmation)

    try:
        await _safe_edit_text(
            callback.message,
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
    if not user:
        await callback.answer('Profil topilmadi. /start ni bosib qayta urinib ko‘ring.', show_alert=True)
        return

    if action == 'cancel':
        await _show_main_menu_from_callback(
            callback,
            state=state,
            users_repo=users_repo,
            answer_text='Jarayon bekor qilindi.',
        )
        return

    current_state = await state.get_state()
    if current_state != CreatePresentationStates.waiting_confirmation.state:
        await callback.answer()
        return

    existing_job, ahead_count = await generation_queue_service.describe_existing_job(callback.from_user.id)
    if existing_job:
        await state.clear()
        await _safe_edit_text(
            callback.message,
            text=create_already_queued_text(ahead_count),
            reply_markup=main_menu_keyboard(is_admin=bool(user.get('is_admin'))),
        )
        await callback.answer('Sizda bitta taqdimot tayyorlanmoqda.', show_alert=True)
        return

    if user.get('generation_access_blocked'):
        await state.clear()
        await _safe_edit_text(
            callback.message,
            text=create_generation_blocked_text(),
            reply_markup=create_credit_missing_keyboard(back_only=True),
        )
        await callback.answer('Slayd yaratish hozircha yopiq.', show_alert=True)
        return

    if not user or not generation_access_service.has_available_generation(user):
        await state.clear()
        await _safe_edit_text(
            callback.message,
            text=create_credit_missing_text(),
            reply_markup=create_credit_missing_keyboard(),
        )
        await callback.answer('Yaratish imkoniyati qolmagan.', show_alert=True)
        return

    consumed_from = await generation_access_service.consume_generation(users_repo, callback.from_user.id)
    if not consumed_from:
        await state.clear()
        await _safe_edit_text(
            callback.message,
            text=create_credit_missing_text(),
            reply_markup=create_credit_missing_keyboard(),
        )
        await callback.answer('Yaratish imkoniyati topilmadi.', show_alert=True)
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
        await _safe_edit_text(
            callback.message,
            text=create_generation_failed_text(),
            reply_markup=main_menu_keyboard(is_admin=bool(user.get('is_admin'))),
        )
        await callback.answer('Hozircha so‘rovni qabul qilib bo‘lmadi.', show_alert=True)
        return

    if active_job:
        await generation_access_service.restore_consumed_generation(users_repo, callback.from_user.id, consumed_from)
        await state.clear()
        await _safe_edit_text(
            callback.message,
            text=create_already_queued_text(ahead_count),
            reply_markup=main_menu_keyboard(is_admin=bool(user.get('is_admin'))),
        )
        await callback.answer('Sizda bitta taqdimot tayyorlanmoqda.', show_alert=True)
        return

    await state.clear()

    try:
        await _safe_edit_text(
            callback.message,
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
    await callback.answer('So‘rovingiz qabul qilindi.')


@router.callback_query(CreateFlowCallback.filter(F.action == 'cancel'))
async def create_cancel_handler(callback: CallbackQuery, state: FSMContext, users_repo: UsersRepository) -> None:
    await _show_main_menu_from_callback(
        callback,
        state=state,
        users_repo=users_repo,
        answer_text='Jarayon bekor qilindi.',
    )
