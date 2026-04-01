from __future__ import annotations

from html import escape

from aiogram import F, Router
from aiogram.filters import Command
from aiogram.fsm.context import FSMContext
from aiogram.types import CallbackQuery, FSInputFile, Message

from app.callbacks.admin import AdminBroadcastCallback, AdminChannelCallback, AdminMenuCallback, AdminUserCallback
from app.filters.admin import AdminFilter
from app.keyboards.admin import (
    admin_broadcast_audience_keyboard,
    admin_broadcast_preview_keyboard,
    admin_broadcast_skip_buttons_keyboard,
    admin_channel_card_keyboard,
    admin_channels_keyboard,
    admin_export_format_keyboard,
    admin_export_keyboard,
    admin_main_keyboard,
    admin_search_results_keyboard,
    admin_secondary_keyboard,
    admin_user_card_keyboard,
)
from app.services.admin import AdminService
from app.services.data_migration import (
    LegacyMongoToCurrentDbMigrationService,
    MigrationBlockedError,
    MigrationConfigError,
)
from app.states.admin import AdminBroadcastStates, AdminChannelStates, AdminUserSearchStates
from app.texts.admin import (
    admin_broadcast_buttons_prompt_text,
    admin_broadcast_content_prompt_text,
    admin_broadcast_menu_text,
    admin_broadcast_preview_text,
    admin_broadcast_result_text,
    admin_channel_add_prompt_text,
    admin_channel_card_text,
    admin_channel_private_link_prompt_text,
    admin_channels_text,
    admin_credit_prompt_text,
    admin_export_ready_text,
    admin_export_text,
    admin_main_menu_text,
    admin_rating_text,
    admin_search_results_text,
    admin_simple_result_text,
    admin_stats_text,
    admin_user_card_text,
    admin_user_search_prompt_text,
)

router = Router(name='admin-panel')
router.message.filter(AdminFilter())
router.callback_query.filter(AdminFilter())


def _migration_summary_text(summary: dict) -> str:
    users = summary.get('users', {})
    referrals = summary.get('referrals', {})
    channels = summary.get('mandatory_channels', {})
    generations = summary.get('generations', {})
    legacy_db_name = escape(str(summary.get('legacy_db_name') or 'legacy'))
    return (
        "<b>✅ Legacy data migratsiyasi yakunlandi</b>\n\n"
        f"• Legacy DB: <code>{legacy_db_name}</code>\n\n"
        "<u>Users</u>\n"
        f"• Insert: <b>{users.get('inserted', 0)}</b>\n"
        f"• Merge: <b>{users.get('merged', 0)}</b>\n"
        f"• Skip: <b>{users.get('skipped', 0)}</b>\n\n"
        "<u>Referrals</u>\n"
        f"• Insert: <b>{referrals.get('inserted', 0)}</b>\n"
        f"• Update: <b>{referrals.get('updated', 0)}</b>\n"
        f"• Skip: <b>{referrals.get('skipped', 0)}</b>\n\n"
        "<u>Mandatory Channels</u>\n"
        f"• Insert: <b>{channels.get('inserted', 0)}</b>\n"
        f"• Update: <b>{channels.get('updated', 0)}</b>\n"
        f"• Skip: <b>{channels.get('skipped', 0)}</b>\n\n"
        "<u>Generations</u>\n"
        f"• Import: <b>{generations.get('inserted', 0)}</b>\n"
        f"• Skip: <b>{generations.get('skipped', 0)}</b>"
    )


async def _show_admin_home(target: Message | CallbackQuery, admin_name: str, *, edit: bool = True) -> None:
    text = admin_main_menu_text(admin_name)
    if isinstance(target, CallbackQuery) and target.message and edit:
        await target.message.edit_text(text=text, reply_markup=admin_main_keyboard())
        await target.answer()
        return

    if isinstance(target, CallbackQuery):
        await target.message.answer(text=text, reply_markup=admin_main_keyboard())
        await target.answer()
        return

    await target.answer(text=text, reply_markup=admin_main_keyboard())


async def _send_broadcast_preview(message: Message, state: FSMContext, admin_service: AdminService) -> None:
    data = await state.get_data()
    draft = data.get('broadcast_draft')
    buttons = data.get('broadcast_buttons') or []
    filter_key = data.get('broadcast_filter', 'all')
    audience_label = admin_service.audience_label(filter_key)
    audience_count = await admin_service.count_audience(filter_key)

    if not draft:
        await message.answer('Preview uchun kontent topilmadi.')
        return

    await admin_service.send_draft(bot=message.bot, chat_id=message.chat.id, draft=draft, buttons=buttons)
    await message.answer(
        text=admin_broadcast_preview_text(audience_label, audience_count, len(buttons), draft),
        reply_markup=admin_broadcast_preview_keyboard(),
    )


async def _open_user_card(target: CallbackQuery | Message, admin_service: AdminService, telegram_id: int) -> None:
    card = await admin_service.build_user_card(telegram_id)
    if not card:
        if isinstance(target, CallbackQuery):
            await target.answer('Foydalanuvchi topilmadi.', show_alert=True)
        else:
            await target.answer('Foydalanuvchi topilmadi.')
        return

    text = admin_user_card_text(card)
    reply_markup = admin_user_card_keyboard(card['user'])

    if isinstance(target, CallbackQuery) and target.message:
        await target.message.edit_text(text=text, reply_markup=reply_markup)
        await target.answer()
        return

    await target.answer(text=text, reply_markup=reply_markup)


async def _show_channels(target: Message | CallbackQuery, admin_service: AdminService, *, edit: bool = True) -> None:
    channels = await admin_service.list_mandatory_channels()
    text = admin_channels_text(channels)
    reply_markup = admin_channels_keyboard(channels)

    if isinstance(target, CallbackQuery) and target.message and edit:
        await target.message.edit_text(text=text, reply_markup=reply_markup)
        await target.answer()
        return

    if isinstance(target, CallbackQuery):
        await target.message.answer(text=text, reply_markup=reply_markup)
        await target.answer()
        return

    await target.answer(text=text, reply_markup=reply_markup)


async def _open_channel_card(target: CallbackQuery | Message, admin_service: AdminService, chat_id: int) -> None:
    channel = await admin_service.get_mandatory_channel(chat_id)
    if not channel:
        if isinstance(target, CallbackQuery):
            await target.answer('Kanal topilmadi.', show_alert=True)
        else:
            await target.answer('Kanal topilmadi.')
        return

    text = admin_channel_card_text(channel)
    reply_markup = admin_channel_card_keyboard(channel)

    if isinstance(target, CallbackQuery) and target.message:
        await target.message.edit_text(text=text, reply_markup=reply_markup)
        await target.answer()
        return

    await target.answer(text=text, reply_markup=reply_markup)


@router.message(Command('admin'))
async def admin_command_handler(message: Message, state: FSMContext) -> None:
    await state.clear()
    await _show_admin_home(message, message.from_user.full_name, edit=False)


@router.message(Command('data_mongodb_to_current_db'))
async def admin_data_mongodb_to_current_db_handler(
    message: Message,
    state: FSMContext,
    data_migration_service: LegacyMongoToCurrentDbMigrationService,
) -> None:
    await state.clear()
    await message.answer('⏳ Legacy MongoDB ma’lumotlari current DB ga ko‘chirilmoqda. Bu biroz vaqt olishi mumkin.')
    try:
        summary = await data_migration_service.run(admin_id=message.from_user.id)
    except MigrationBlockedError as exc:
        await message.answer(admin_simple_result_text(str(exc)))
        return
    except MigrationConfigError as exc:
        await message.answer(admin_simple_result_text(str(exc)))
        return
    except Exception as exc:
        await message.answer(
            admin_simple_result_text(
                f"Migratsiya yakunlanmadi: {escape(str(exc) or 'noma’lum xatolik')}"
            )
        )
        return

    await message.answer(_migration_summary_text(summary))


@router.callback_query(AdminMenuCallback.filter(F.action == 'main'))
async def admin_main_handler(callback: CallbackQuery, state: FSMContext) -> None:
    await state.clear()
    await _show_admin_home(callback, callback.from_user.full_name)


@router.callback_query(AdminMenuCallback.filter(F.action == 'stats'))
async def admin_stats_handler(callback: CallbackQuery, admin_service: AdminService, state: FSMContext) -> None:
    await state.clear()
    stats = await admin_service.build_statistics()
    await callback.message.edit_text(
        text=admin_stats_text(stats),
        reply_markup=admin_secondary_keyboard(),
    )
    await callback.answer()


@router.callback_query(AdminMenuCallback.filter(F.action == 'rating'))
async def admin_rating_handler(callback: CallbackQuery, admin_service: AdminService, state: FSMContext) -> None:
    await state.clear()
    ratings = await admin_service.build_ratings()
    await callback.message.edit_text(
        text=admin_rating_text(ratings),
        reply_markup=admin_secondary_keyboard(),
    )
    await callback.answer()


@router.callback_query(AdminMenuCallback.filter(F.action == 'users'))
async def admin_users_handler(callback: CallbackQuery, state: FSMContext) -> None:
    await state.clear()
    await state.set_state(AdminUserSearchStates.waiting_query)
    await callback.message.edit_text(
        text=admin_user_search_prompt_text(),
        reply_markup=admin_secondary_keyboard(),
    )
    await callback.answer()


@router.callback_query(AdminMenuCallback.filter(F.action == 'channels'))
async def admin_channels_handler(callback: CallbackQuery, state: FSMContext, admin_service: AdminService) -> None:
    await state.clear()
    await _show_channels(callback, admin_service)


@router.callback_query(AdminChannelCallback.filter(F.action == 'add'))
async def admin_channel_add_handler(callback: CallbackQuery, state: FSMContext) -> None:
    await state.clear()
    await state.set_state(AdminChannelStates.waiting_channel_reference)
    await callback.message.edit_text(
        text=admin_channel_add_prompt_text(),
        reply_markup=admin_secondary_keyboard(),
    )
    await callback.answer()


@router.message(AdminChannelStates.waiting_channel_reference)
async def admin_channel_reference_handler(message: Message, state: FSMContext, admin_service: AdminService) -> None:
    raw_reference = (message.text or '').strip()
    try:
        channel_payload = await admin_service.resolve_channel_reference(bot=message.bot, raw_reference=raw_reference)
    except ValueError as exc:
        await message.answer(str(exc))
        return

    if channel_payload.get('invite_link'):
        await admin_service.save_mandatory_channel(channel_payload)
        await state.clear()
        await message.answer(admin_simple_result_text('Kanal muvaffaqiyatli saqlandi.'))
        await _show_channels(message, admin_service, edit=False)
        return

    await state.set_state(AdminChannelStates.waiting_channel_invite_link)
    await state.update_data(pending_channel_payload=channel_payload)
    await message.answer(
        text=admin_channel_private_link_prompt_text(channel_payload),
        reply_markup=admin_secondary_keyboard(),
    )


@router.message(AdminChannelStates.waiting_channel_invite_link)
async def admin_channel_invite_link_handler(message: Message, state: FSMContext, admin_service: AdminService) -> None:
    data = await state.get_data()
    channel_payload = data.get('pending_channel_payload')
    if not channel_payload:
        await state.clear()
        await message.answer('Kanal ma’lumoti topilmadi. Qaytadan urinib ko‘ring.')
        return

    try:
        channel_payload['invite_link'] = admin_service.normalize_invite_link(message.text or '')
    except ValueError as exc:
        await message.answer(str(exc))
        return

    await admin_service.save_mandatory_channel(channel_payload)
    await state.clear()
    await message.answer(admin_simple_result_text('Private kanal muvaffaqiyatli saqlandi.'))
    await _show_channels(message, admin_service, edit=False)


@router.message(AdminUserSearchStates.waiting_query)
async def admin_user_search_message_handler(message: Message, state: FSMContext, admin_service: AdminService) -> None:
    query = (message.text or '').strip()
    if len(query) < 2:
        await message.answer('Qidiruv so‘rovi kamida 2 ta belgidan iborat bo‘lishi kerak.')
        return

    results = await admin_service.search_users(query)
    await state.update_data(last_search_query=query)
    await message.answer(
        text=admin_search_results_text(query, results),
        reply_markup=admin_search_results_keyboard(results),
    )


@router.callback_query(AdminUserCallback.filter(F.action == 'open'))
async def admin_open_user_handler(callback: CallbackQuery, callback_data: AdminUserCallback, admin_service: AdminService, state: FSMContext) -> None:
    await state.clear()
    await _open_user_card(callback, admin_service, callback_data.user_id)


@router.callback_query(AdminChannelCallback.filter(F.action == 'open'))
async def admin_channel_open_handler(callback: CallbackQuery, callback_data: AdminChannelCallback, admin_service: AdminService) -> None:
    await _open_channel_card(callback, admin_service, callback_data.chat_id)


@router.callback_query(AdminChannelCallback.filter(F.action == 'toggle'))
async def admin_channel_toggle_handler(callback: CallbackQuery, callback_data: AdminChannelCallback, admin_service: AdminService) -> None:
    channel = await admin_service.get_mandatory_channel(callback_data.chat_id)
    if not channel:
        await callback.answer('Kanal topilmadi.', show_alert=True)
        return

    await admin_service.set_channel_active(callback_data.chat_id, not bool(channel.get('is_active')))
    await _open_channel_card(callback, admin_service, callback_data.chat_id)


@router.callback_query(AdminChannelCallback.filter(F.action == 'delete'))
async def admin_channel_delete_handler(callback: CallbackQuery, callback_data: AdminChannelCallback, admin_service: AdminService) -> None:
    deleted = await admin_service.delete_channel(callback_data.chat_id)
    if not deleted:
        await callback.answer('Kanal topilmadi.', show_alert=True)
        return

    await callback.answer('Kanal o‘chirildi.', show_alert=True)
    await _show_channels(callback, admin_service)


@router.callback_query(AdminUserCallback.filter(F.action.in_({'toggle_unlimited', 'toggle_generation_block', 'toggle_bot_block'})))
async def admin_user_toggle_handler(
    callback: CallbackQuery,
    callback_data: AdminUserCallback,
    admin_service: AdminService,
    users_repo,
) -> None:
    user = await users_repo.get_by_telegram_id(callback_data.user_id)
    if not user:
        await callback.answer('Foydalanuvchi topilmadi.', show_alert=True)
        return

    if callback_data.action == 'toggle_unlimited':
        await users_repo.set_generation_unlimited(callback_data.user_id, not bool(user.get('generation_unlimited')))
    elif callback_data.action == 'toggle_generation_block':
        await users_repo.set_generation_access_blocked(callback_data.user_id, not bool(user.get('generation_access_blocked')))
    elif callback_data.action == 'toggle_bot_block':
        await users_repo.set_bot_access_blocked(callback_data.user_id, not bool(user.get('bot_access_blocked')))

    await _open_user_card(callback, admin_service, callback_data.user_id)


@router.callback_query(AdminUserCallback.filter(F.action.in_({'credit_add', 'credit_remove'})))
async def admin_user_credit_prompt_handler(
    callback: CallbackQuery,
    callback_data: AdminUserCallback,
    users_repo,
    state: FSMContext,
) -> None:
    user = await users_repo.get_by_telegram_id(callback_data.user_id)
    if not user:
        await callback.answer('Foydalanuvchi topilmadi.', show_alert=True)
        return

    operation = 'add' if callback_data.action == 'credit_add' else 'remove'
    await state.set_state(AdminUserSearchStates.waiting_credit_amount)
    await state.update_data(target_user_id=callback_data.user_id, credit_operation=operation)
    await callback.message.answer(
        text=admin_credit_prompt_text(user, operation),
        reply_markup=admin_secondary_keyboard(),
    )
    await callback.answer()


@router.message(AdminUserSearchStates.waiting_credit_amount)
async def admin_user_credit_amount_handler(message: Message, state: FSMContext, users_repo, admin_service: AdminService) -> None:
    data = await state.get_data()
    target_user_id = int(data.get('target_user_id', 0) or 0)
    operation = data.get('credit_operation', 'add')

    if not target_user_id:
        await state.clear()
        await message.answer('Foydalanuvchi tanlanmagan. Qaytadan urinib ko‘ring.')
        return

    try:
        amount = int((message.text or '').strip())
    except Exception:
        await message.answer('Butun son kiriting. Masalan: 2')
        return

    if amount <= 0:
        await message.answer('Qiymat musbat butun son bo‘lishi kerak.')
        return

    delta = amount if operation == 'add' else -amount
    updated = await users_repo.adjust_bonus_generation_credits(target_user_id, delta)
    await state.clear()

    if not updated:
        await message.answer('Foydalanuvchi topilmadi.')
        return

    await message.answer(admin_simple_result_text('Kreditlar muvaffaqiyatli yangilandi.'))
    await _open_user_card(message, admin_service, target_user_id)


@router.callback_query(AdminMenuCallback.filter(F.action == 'broadcast'))
async def admin_broadcast_menu_handler(callback: CallbackQuery, state: FSMContext) -> None:
    await state.clear()
    await callback.message.edit_text(
        text=admin_broadcast_menu_text(),
        reply_markup=admin_broadcast_audience_keyboard(),
    )
    await callback.answer()


@router.callback_query(AdminBroadcastCallback.filter(F.action == 'audience'))
async def admin_broadcast_audience_handler(
    callback: CallbackQuery,
    callback_data: AdminBroadcastCallback,
    state: FSMContext,
    admin_service: AdminService,
) -> None:
    filter_key = callback_data.value
    audience_count = await admin_service.count_audience(filter_key)
    await state.clear()
    await state.set_state(AdminBroadcastStates.waiting_content)
    await state.update_data(broadcast_filter=filter_key)
    await callback.message.edit_text(
        text=admin_broadcast_content_prompt_text(admin_service.audience_label(filter_key), audience_count),
        reply_markup=admin_secondary_keyboard(),
    )
    await callback.answer()


@router.message(AdminBroadcastStates.waiting_content)
async def admin_broadcast_content_handler(message: Message, state: FSMContext, admin_service: AdminService) -> None:
    draft = admin_service.extract_draft_from_message(message)
    if not draft:
        await message.answer('Faqat matn, rasm, video, GIF yoki document yuboring.')
        return

    await state.update_data(broadcast_draft=draft, broadcast_buttons=[])
    await state.set_state(AdminBroadcastStates.waiting_buttons)
    await message.answer(
        text=admin_broadcast_buttons_prompt_text(),
        reply_markup=admin_broadcast_skip_buttons_keyboard(),
    )


@router.callback_query(AdminBroadcastCallback.filter(F.action == 'buttons_skip'))
async def admin_broadcast_skip_buttons_handler(callback: CallbackQuery, state: FSMContext, admin_service: AdminService) -> None:
    await state.update_data(broadcast_buttons=[])
    await _send_broadcast_preview(callback.message, state, admin_service)
    await callback.answer()


@router.message(AdminBroadcastStates.waiting_buttons)
async def admin_broadcast_buttons_handler(message: Message, state: FSMContext, admin_service: AdminService) -> None:
    try:
        buttons = admin_service.parse_buttons(message.text or '')
    except ValueError as exc:
        await message.answer(str(exc))
        return

    await state.update_data(broadcast_buttons=buttons)
    await _send_broadcast_preview(message, state, admin_service)


@router.callback_query(AdminBroadcastCallback.filter(F.action == 'edit_content'))
async def admin_broadcast_edit_content_handler(callback: CallbackQuery, state: FSMContext, admin_service: AdminService) -> None:
    data = await state.get_data()
    filter_key = data.get('broadcast_filter', 'all')
    await state.set_state(AdminBroadcastStates.waiting_content)
    await callback.message.edit_text(
        text=admin_broadcast_content_prompt_text(
            admin_service.audience_label(filter_key),
            await admin_service.count_audience(filter_key),
        ),
        reply_markup=admin_secondary_keyboard(),
    )
    await callback.answer()


@router.callback_query(AdminBroadcastCallback.filter(F.action == 'edit_buttons'))
async def admin_broadcast_edit_buttons_handler(callback: CallbackQuery, state: FSMContext) -> None:
    await state.set_state(AdminBroadcastStates.waiting_buttons)
    await callback.message.edit_text(
        text=admin_broadcast_buttons_prompt_text(),
        reply_markup=admin_broadcast_skip_buttons_keyboard(),
    )
    await callback.answer()


@router.callback_query(AdminBroadcastCallback.filter(F.action == 'test'))
async def admin_broadcast_test_handler(callback: CallbackQuery, state: FSMContext, admin_service: AdminService) -> None:
    data = await state.get_data()
    draft = data.get('broadcast_draft')
    buttons = data.get('broadcast_buttons') or []
    if not draft:
        await callback.answer('Xabar kontenti topilmadi.', show_alert=True)
        return

    await admin_service.send_draft(bot=callback.bot, chat_id=callback.from_user.id, draft=draft, buttons=buttons)
    await callback.answer('Test xabar yuborildi.', show_alert=True)


@router.callback_query(AdminBroadcastCallback.filter(F.action == 'send'))
async def admin_broadcast_send_handler(callback: CallbackQuery, state: FSMContext, admin_service: AdminService) -> None:
    data = await state.get_data()
    draft = data.get('broadcast_draft')
    buttons = data.get('broadcast_buttons') or []
    filter_key = data.get('broadcast_filter', 'all')
    if not draft:
        await callback.answer('Xabar kontenti topilmadi.', show_alert=True)
        return

    await callback.message.edit_text('⏳ Ommaviy xabar yuborilmoqda. Iltimos, natijani kuting...')
    result = await admin_service.broadcast(bot=callback.bot, filter_key=filter_key, draft=draft, buttons=buttons)
    await state.clear()
    await callback.message.edit_text(
        text=admin_broadcast_result_text(admin_service.audience_label(filter_key), result),
        reply_markup=admin_main_keyboard(),
    )
    await callback.answer('Yuborish yakunlandi.')


@router.callback_query(AdminBroadcastCallback.filter(F.action == 'cancel'))
async def admin_broadcast_cancel_handler(callback: CallbackQuery, state: FSMContext) -> None:
    await state.clear()
    await callback.message.edit_text(
        text=admin_main_menu_text(callback.from_user.full_name),
        reply_markup=admin_main_keyboard(),
    )
    await callback.answer('Jarayon bekor qilindi.')


@router.callback_query(AdminMenuCallback.filter(F.action == 'exports'))
async def admin_exports_handler(callback: CallbackQuery, state: FSMContext) -> None:
    await state.clear()
    await callback.message.edit_text(
        text=admin_export_text(),
        reply_markup=admin_export_keyboard(),
    )
    await callback.answer()


@router.callback_query(AdminMenuCallback.filter(F.action == 'export_audience'))
async def admin_export_audience_handler(callback: CallbackQuery, callback_data: AdminMenuCallback) -> None:
    await callback.message.edit_text(
        text='Kerakli eksport formatini tanlang.',
        reply_markup=admin_export_format_keyboard(callback_data.value or 'all'),
    )
    await callback.answer()


@router.callback_query(AdminMenuCallback.filter(F.action == 'export_format'))
async def admin_export_format_handler(callback: CallbackQuery, callback_data: AdminMenuCallback, admin_service: AdminService) -> None:
    raw_value = callback_data.value or 'all__csv'
    parts = raw_value.rsplit('__', 1)
    if len(parts) != 2:
        filter_key, fmt = 'all', 'csv'
    else:
        filter_key, fmt = parts
    path, count = await admin_service.export_users(filter_key=filter_key, fmt=fmt)
    try:
        await callback.message.answer_document(
            document=FSInputFile(path),
            caption=admin_export_ready_text(admin_service.audience_label(filter_key), fmt, count),
        )
    finally:
        admin_service.cleanup_file(path)

    await callback.answer('Eksport fayli tayyor.')
