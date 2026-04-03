from __future__ import annotations

from aiogram import F, Router
from aiogram.fsm.context import FSMContext
from aiogram.types import CallbackQuery, Message

from app.callbacks.admin import AdminMenuCallback
from app.callbacks.magic import MagicAdminCallback, MagicCardCallback, MagicTopupCallback
from app.filters.admin import AdminFilter
from app.handlers.user.magic import notify_magic_topup_approved, notify_magic_topup_rejected
from app.keyboards.admin import admin_secondary_keyboard
from app.keyboards.magic import (
    admin_magic_card_keyboard,
    admin_magic_cards_keyboard,
    admin_magic_pending_keyboard,
    admin_magic_settings_keyboard,
    admin_magic_topup_review_keyboard,
)
from app.services.magic_slides import MagicSlideService
from app.states.magic import AdminMagicStates
from app.texts.magic import (
    magic_admin_card_prompt_text,
    magic_admin_card_text,
    magic_admin_cards_text,
    magic_admin_pending_text,
    magic_admin_price_prompt_text,
    magic_admin_settings_text,
)

router = Router(name='admin-magic')
router.message.filter(AdminFilter())
router.callback_query.filter(AdminFilter())


async def _show_magic_settings(target: Message | CallbackQuery, magic_slide_service: MagicSlideService, *, edit: bool = True) -> None:
    context = await magic_slide_service.get_settings_context()
    text = magic_admin_settings_text(context)
    reply_markup = admin_magic_settings_keyboard(maintenance_enabled=bool(context['settings'].get('maintenance_enabled')))

    if isinstance(target, CallbackQuery) and target.message and edit:
        await target.message.edit_text(text=text, reply_markup=reply_markup)
        await target.answer()
        return

    if isinstance(target, CallbackQuery):
        await target.message.answer(text=text, reply_markup=reply_markup)
        await target.answer()
        return

    await target.answer(text=text, reply_markup=reply_markup)


@router.callback_query(AdminMenuCallback.filter(F.action == 'magic'))
async def admin_magic_settings_handler(callback: CallbackQuery, state: FSMContext, magic_slide_service: MagicSlideService) -> None:
    await state.clear()
    await _show_magic_settings(callback, magic_slide_service)


@router.callback_query(MagicAdminCallback.filter(F.action == 'settings'))
async def admin_magic_settings_refresh_handler(callback: CallbackQuery, state: FSMContext, magic_slide_service: MagicSlideService) -> None:
    await state.clear()
    await _show_magic_settings(callback, magic_slide_service)


@router.callback_query(MagicAdminCallback.filter(F.action == 'price'))
async def admin_magic_price_prompt_handler(callback: CallbackQuery, state: FSMContext, magic_slide_service: MagicSlideService) -> None:
    await state.clear()
    context = await magic_slide_service.get_settings_context()
    await state.set_state(AdminMagicStates.waiting_price)
    await callback.message.edit_text(
        text=magic_admin_price_prompt_text(int(context['settings'].get('price_per_presentation', 0) or 0)),
        reply_markup=admin_secondary_keyboard(),
    )
    await callback.answer()


@router.message(AdminMagicStates.waiting_price)
async def admin_magic_price_value_handler(message: Message, state: FSMContext, magic_slide_service: MagicSlideService) -> None:
    raw_value = ''.join(ch for ch in (message.text or '') if ch.isdigit())
    if not raw_value:
        await message.answer('Narxni butun son ko‘rinishida yuboring. Masalan: 15000')
        return

    try:
        await magic_slide_service.set_price(int(raw_value))
    except ValueError as exc:
        await message.answer(str(exc))
        return

    await state.clear()
    await message.answer('✅ Magic Slayd narxi yangilandi.')
    await _show_magic_settings(message, magic_slide_service, edit=False)


@router.callback_query(MagicAdminCallback.filter(F.action == 'maintenance_toggle'))
async def admin_magic_maintenance_toggle_handler(callback: CallbackQuery, magic_slide_service: MagicSlideService) -> None:
    await magic_slide_service.toggle_maintenance()
    await _show_magic_settings(callback, magic_slide_service)


@router.callback_query(MagicAdminCallback.filter(F.action == 'cards'))
async def admin_magic_cards_handler(callback: CallbackQuery, state: FSMContext, magic_slide_service: MagicSlideService) -> None:
    await state.clear()
    cards = await magic_slide_service.list_cards()
    cards_view = [{**card, 'masked_number': magic_slide_service.mask_card_number(card.get('card_number'))} for card in cards]
    await callback.message.edit_text(
        text=magic_admin_cards_text(cards_view),
        reply_markup=admin_magic_cards_keyboard(cards_view),
    )
    await callback.answer()


@router.callback_query(MagicAdminCallback.filter(F.action == 'cards_add'))
async def admin_magic_card_add_prompt_handler(callback: CallbackQuery, state: FSMContext) -> None:
    await state.clear()
    await state.set_state(AdminMagicStates.waiting_card_details)
    await callback.message.edit_text(
        text=magic_admin_card_prompt_text(),
        reply_markup=admin_secondary_keyboard(),
    )
    await callback.answer()


@router.message(AdminMagicStates.waiting_card_details)
async def admin_magic_card_add_value_handler(message: Message, state: FSMContext, magic_slide_service: MagicSlideService) -> None:
    try:
        await magic_slide_service.create_card(message.text or '')
    except ValueError as exc:
        await message.answer(str(exc))
        return

    await state.clear()
    await message.answer('✅ Yangi karta qo‘shildi.')
    await _show_magic_settings(message, magic_slide_service, edit=False)


@router.callback_query(MagicCardCallback.filter(F.action == 'open'))
async def admin_magic_card_open_handler(callback: CallbackQuery, callback_data: MagicCardCallback, magic_slide_service: MagicSlideService) -> None:
    card = await magic_slide_service.get_card(callback_data.card_id)
    if not card:
        await callback.answer('Karta topilmadi.', show_alert=True)
        return

    card_view = {**card, 'masked_number': magic_slide_service.mask_card_number(card.get('card_number'))}
    await callback.message.edit_text(
        text=magic_admin_card_text(card_view),
        reply_markup=admin_magic_card_keyboard(card),
    )
    await callback.answer()


@router.callback_query(MagicCardCallback.filter(F.action == 'toggle'))
async def admin_magic_card_toggle_handler(callback: CallbackQuery, callback_data: MagicCardCallback, magic_slide_service: MagicSlideService) -> None:
    card = await magic_slide_service.toggle_card(callback_data.card_id)
    if not card:
        await callback.answer('Karta topilmadi.', show_alert=True)
        return

    card_view = {**card, 'masked_number': magic_slide_service.mask_card_number(card.get('card_number'))}
    await callback.message.edit_text(
        text=magic_admin_card_text(card_view),
        reply_markup=admin_magic_card_keyboard(card),
    )
    await callback.answer()


@router.callback_query(MagicCardCallback.filter(F.action == 'delete'))
async def admin_magic_card_delete_handler(callback: CallbackQuery, callback_data: MagicCardCallback, magic_slide_service: MagicSlideService) -> None:
    deleted = await magic_slide_service.delete_card(callback_data.card_id)
    if not deleted:
        await callback.answer('Karta topilmadi.', show_alert=True)
        return

    await callback.answer('Karta o‘chirildi.', show_alert=True)
    await _show_magic_settings(callback, magic_slide_service)


@router.callback_query(MagicAdminCallback.filter(F.action == 'pending'))
async def admin_magic_pending_handler(callback: CallbackQuery, state: FSMContext, magic_slide_service: MagicSlideService) -> None:
    await state.clear()
    topups = await magic_slide_service.list_pending_topups()
    await callback.message.edit_text(
        text=magic_admin_pending_text(topups),
        reply_markup=admin_magic_pending_keyboard(topups),
    )
    await callback.answer()


@router.callback_query(MagicTopupCallback.filter(F.action == 'open'))
async def admin_magic_pending_open_handler(callback: CallbackQuery, callback_data: MagicTopupCallback, magic_slide_service: MagicSlideService) -> None:
    topup = await magic_slide_service.get_topup(callback_data.topup_id)
    if not topup:
        await callback.answer('To‘lov so‘rovi topilmadi.', show_alert=True)
        return
    if topup.get('status') != 'pending':
        await callback.answer('Bu so‘rov allaqachon ko‘rib chiqilgan.', show_alert=True)
        return

    await magic_slide_service.resend_topup_to_admin(
        bot=callback.bot,
        chat_id=callback.from_user.id,
        topup_id=callback_data.topup_id,
        reply_markup=admin_magic_topup_review_keyboard(callback_data.topup_id),
    )
    await callback.answer('Chek sizga qayta yuborildi.', show_alert=True)


@router.callback_query(MagicTopupCallback.filter(F.action == 'approve'))
async def admin_magic_topup_approve_handler(callback: CallbackQuery, callback_data: MagicTopupCallback, magic_slide_service: MagicSlideService) -> None:
    topup = await magic_slide_service.approve_topup(
        topup_id=callback_data.topup_id,
        admin_id=callback.from_user.id,
        admin_name=callback.from_user.full_name,
    )
    if not topup:
        await callback.answer('Bu so‘rov allaqachon ko‘rib chiqilgan.', show_alert=True)
        if callback.message:
            await callback.message.edit_reply_markup(reply_markup=None)
        return

    await magic_slide_service.clear_admin_review_keyboards(callback.bot, topup)
    try:
        await notify_magic_topup_approved(
            callback.bot,
            int(topup['telegram_id']),
            int(topup['amount_uzs']),
            int(topup.get('account_balance_uzs', 0) or 0),
        )
    except Exception:
        pass

    await callback.answer('To‘lov tasdiqlandi.', show_alert=True)


@router.callback_query(MagicTopupCallback.filter(F.action == 'reject'))
async def admin_magic_topup_reject_handler(callback: CallbackQuery, callback_data: MagicTopupCallback, magic_slide_service: MagicSlideService) -> None:
    topup = await magic_slide_service.reject_topup(
        topup_id=callback_data.topup_id,
        admin_id=callback.from_user.id,
        admin_name=callback.from_user.full_name,
    )
    if not topup:
        await callback.answer('Bu so‘rov allaqachon ko‘rib chiqilgan.', show_alert=True)
        if callback.message:
            await callback.message.edit_reply_markup(reply_markup=None)
        return

    await magic_slide_service.clear_admin_review_keyboards(callback.bot, topup)
    try:
        await notify_magic_topup_rejected(
            callback.bot,
            int(topup['telegram_id']),
            int(topup['amount_uzs']),
        )
    except Exception:
        pass

    await callback.answer('To‘lov rad etildi.', show_alert=True)
