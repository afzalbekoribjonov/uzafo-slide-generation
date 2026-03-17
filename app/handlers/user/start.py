from __future__ import annotations

from aiogram import Router
from aiogram.filters import CommandStart
from aiogram.filters.command import CommandObject
from aiogram.types import Message

from app.keyboards.user import main_menu_keyboard, subscription_keyboard
from app.services.referrals import ReferralService
from app.services.subscriptions import SubscriptionService
from app.services.users import UserService
from app.texts.user import main_menu_text, subscription_failed_text
from app.utils.deeplink import parse_inviter_id

router = Router(name='user-start')


@router.message(CommandStart())
async def start_handler(
    message: Message,
    command: CommandObject,
    user_service: UserService,
    referral_service: ReferralService,
    subscription_service: SubscriptionService,
) -> None:
    inviter_id = parse_inviter_id(command.args if command else None)

    user, is_new_user = await user_service.get_or_create_user(
        message.from_user,
        invited_by=inviter_id,
    )

    await referral_service.register_start_if_valid(
        inviter_id=inviter_id,
        invitee_id=message.from_user.id,
        is_new_user=is_new_user,
    )

    channels = await subscription_service.get_active_channels()

    if channels and not user.get('is_admin'):
        is_subscribed, unsubscribed_channels = await subscription_service.check_user_subscriptions(
            bot=message.bot,
            user_id=message.from_user.id,
        )

        if not is_subscribed:
            await user_service.set_subscription_verified(message.from_user.id, False)
            await message.answer(
                text=subscription_failed_text(unsubscribed_channels),
                reply_markup=subscription_keyboard(channels),
            )
            return

        await user_service.set_subscription_verified(message.from_user.id, True)
        await referral_service.approve_after_subscription(message.from_user.id)

        user = await user_service.get_user(message.from_user.id)
        await message.answer(
            text=main_menu_text(user['full_name']),
            reply_markup=main_menu_keyboard(is_admin=bool(user.get('is_admin'))),
        )
        return

    await user_service.set_subscription_verified(message.from_user.id, True)
    reward_info = await referral_service.approve_after_subscription(message.from_user.id)
    if reward_info:
        try:
            await message.bot.send_message(
                reward_info['inviter_id'],
                (
                    "<b>🎉 Yangi referral tasdiqlandi</b>\n\n"
                    f"👤 Foydalanuvchi: <b>{reward_info['invitee_name']}</b>\n"
                    "✅ Botdan foydalanish shartlarini bajardi\n"
                    "🎁 Sizga 1 ta qo‘shimcha yaratish imkoni qo‘shildi\n\n"
                    f"📊 Jami tasdiqlangan referral: <b>{reward_info['referral_count']}</b>\n"
                    f"🎟 Mavjud referral kreditlar: <b>{reward_info['referral_credits']}</b>"
                ),
            )
        except Exception:
            pass
    user = await user_service.get_user(message.from_user.id)
    await message.answer(
        text=main_menu_text(user['full_name']),
        reply_markup=main_menu_keyboard(is_admin=bool(user.get('is_admin'))),
    )
