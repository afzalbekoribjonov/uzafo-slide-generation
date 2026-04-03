from __future__ import annotations

from aiogram import F, Router
from aiogram.fsm.context import FSMContext
from aiogram.types import CallbackQuery, ReplyKeyboardRemove

from app.callbacks.admin import PublicPostCallback
from app.callbacks.menu import MenuCallback, StatusCallback
from app.keyboards.user import (
    contact_keyboard,
    help_keyboard,
    invite_keyboard,
    main_menu_keyboard,
    referrals_keyboard,
    status_keyboard,
)
from app.repositories.users import UsersRepository
from app.services.generations import GenerationAccessService
from app.services.referrals import ReferralService
from app.texts.user import (
    contact_text,
    help_text,
    invite_text,
    main_menu_text,
    referrals_text,
    status_text,
)

router = Router(name='user-menu')


@router.callback_query(MenuCallback.filter(F.action == 'main'))
async def menu_main_handler(callback: CallbackQuery, users_repo: UsersRepository, state: FSMContext) -> None:
    data = await state.get_data()
    prompt_chat_id = data.get('magic_webapp_prompt_chat_id')
    prompt_message_id = data.get('magic_webapp_prompt_message_id')
    if prompt_chat_id and prompt_message_id:
        try:
            await callback.bot.delete_message(chat_id=int(prompt_chat_id), message_id=int(prompt_message_id))
        except Exception:
            pass
        await callback.message.answer(
            'Magic Slayd yaratish oynasi yopildi.',
            reply_markup=ReplyKeyboardRemove(),
        )
    await state.clear()
    user = await users_repo.get_by_telegram_id(callback.from_user.id)
    await callback.message.edit_text(
        text=main_menu_text(user['full_name']),
        reply_markup=main_menu_keyboard(is_admin=bool(user.get('is_admin'))),
    )
    await callback.answer()


@router.callback_query(MenuCallback.filter(F.action == 'status'))
async def menu_status_handler(
    callback: CallbackQuery,
    users_repo: UsersRepository,
    generation_access_service: GenerationAccessService,
) -> None:
    user = await users_repo.get_by_telegram_id(callback.from_user.id)
    available = generation_access_service.available_generations(user)

    await callback.message.edit_text(
        text=status_text(user, available),
        reply_markup=status_keyboard(),
    )
    await callback.answer()


@router.callback_query(MenuCallback.filter(F.action == 'invite'))
async def menu_invite_handler(
    callback: CallbackQuery,
    users_repo: UsersRepository,
    generation_access_service: GenerationAccessService,
    bot_username: str,
) -> None:
    user = await users_repo.get_by_telegram_id(callback.from_user.id)
    available = generation_access_service.available_generations(user)
    referral_link = f'https://t.me/{bot_username}?start={callback.from_user.id}'

    await callback.message.edit_text(
        text=invite_text(
            referral_link=referral_link,
            available_generations=available,
            referral_count=user.get('referral_count', 0),
            user=user,
        ),
        reply_markup=invite_keyboard(),
    )
    await callback.answer()


@router.callback_query(StatusCallback.filter(F.action == 'referrals'))
async def referrals_handler(
    callback: CallbackQuery,
    referral_service: ReferralService,
) -> None:
    referrals = await referral_service.list_invited_users(callback.from_user.id)
    await callback.message.edit_text(
        text=referrals_text(referrals),
        reply_markup=referrals_keyboard(),
    )
    await callback.answer()


@router.callback_query(MenuCallback.filter(F.action == 'help'))
async def help_handler(callback: CallbackQuery) -> None:
    await callback.message.edit_text(
        text=help_text(),
        reply_markup=help_keyboard(),
    )
    await callback.answer()


@router.callback_query(MenuCallback.filter(F.action == 'contact'))
async def contact_handler(callback: CallbackQuery, support_contact: str) -> None:
    await callback.message.edit_text(
        text=contact_text(support_contact),
        reply_markup=contact_keyboard(),
    )
    await callback.answer()


@router.callback_query(PublicPostCallback.filter())
async def public_post_button_handler(
    callback: CallbackQuery,
    callback_data: PublicPostCallback,
    users_repo: UsersRepository,
    generation_access_service: GenerationAccessService,
    bot_username: str,
    support_contact: str,
) -> None:
    action = callback_data.action
    user = await users_repo.get_by_telegram_id(callback.from_user.id)
    if not user:
        await callback.answer('Foydalanuvchi topilmadi.', show_alert=True)
        return

    if action == 'main':
        await callback.message.answer(
            text=main_menu_text(user['full_name']),
            reply_markup=main_menu_keyboard(is_admin=bool(user.get('is_admin'))),
        )
        await callback.answer()
        return

    if action == 'status':
        available = generation_access_service.available_generations(user)
        await callback.message.answer(
            text=status_text(user, available),
            reply_markup=status_keyboard(),
        )
        await callback.answer()
        return

    if action == 'invite':
        available = generation_access_service.available_generations(user)
        referral_link = f'https://t.me/{bot_username}?start={callback.from_user.id}'
        await callback.message.answer(
            text=invite_text(
                referral_link=referral_link,
                available_generations=available,
                referral_count=user.get('referral_count', 0),
                user=user,
            ),
            reply_markup=invite_keyboard(),
        )
        await callback.answer()
        return

    if action == 'help':
        await callback.message.answer(text=help_text(), reply_markup=help_keyboard())
        await callback.answer()
        return

    if action == 'contact':
        await callback.message.answer(text=contact_text(support_contact), reply_markup=contact_keyboard())
        await callback.answer()
        return

    await callback.answer()
