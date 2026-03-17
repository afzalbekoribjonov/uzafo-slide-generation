from __future__ import annotations

from aiogram import F, Router
from aiogram.exceptions import TelegramBadRequest
from aiogram.types import CallbackQuery

from app.callbacks.subscription import SubscriptionCallback
from app.keyboards.user import main_menu_keyboard, subscription_keyboard
from app.repositories.users import UsersRepository
from app.services.referrals import ReferralService
from app.services.subscriptions import SubscriptionService
from app.texts.user import main_menu_text, subscription_failed_text

router = Router(name='user-subscription')


@router.callback_query(SubscriptionCallback.filter(F.action == 'check'))
async def subscription_check_handler(
    callback: CallbackQuery,
    users_repo: UsersRepository,
    referral_service: ReferralService,
    subscription_service: SubscriptionService,
) -> None:
    is_subscribed, unsubscribed_channels = await subscription_service.check_user_subscriptions(
        callback.bot,
        callback.from_user.id,
    )

    if not is_subscribed:
        await users_repo.set_subscription_verified(callback.from_user.id, False)
        all_channels = await subscription_service.get_active_channels()

        try:
            await callback.message.edit_text(
                text=subscription_failed_text(unsubscribed_channels),
                reply_markup=subscription_keyboard(all_channels),
            )
        except TelegramBadRequest as e:
            if 'message is not modified' not in str(e):
                raise

        await callback.answer('Barcha majburiy kanallarga obuna bo‘ling.', show_alert=True)
        return

    await users_repo.mark_subscription_verified(callback.from_user.id)
    reward_info = await referral_service.approve_after_subscription(callback.from_user.id)

    if reward_info:
        try:
            await callback.bot.send_message(
                reward_info['inviter_id'],
                (
                    "<b>🎉 Yangi referral tasdiqlandi</b>\n\n"
                    f"👤 Foydalanuvchi: <b>{reward_info['invitee_name']}</b>\n"
                    "✅ Barcha majburiy kanallarga obuna bo‘ldi\n"
                    "🎁 Sizga 1 ta qo‘shimcha yaratish imkoni qo‘shildi\n\n"
                    f"📊 Jami tasdiqlangan referral: <b>{reward_info['referral_count']}</b>\n"
                    f"🎟 Mavjud referral kreditlar: <b>{reward_info['referral_credits']}</b>"
                ),
            )
        except Exception:
            pass

    user = await users_repo.get_by_telegram_id(callback.from_user.id)
    await callback.message.edit_text(
        text=main_menu_text(user['full_name']),
        reply_markup=main_menu_keyboard(is_admin=bool(user.get('is_admin'))),
    )
    await callback.answer('Obuna muvaffaqiyatli tasdiqlandi.')
