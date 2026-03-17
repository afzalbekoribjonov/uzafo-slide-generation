from __future__ import annotations

from typing import Any, Awaitable, Callable

from aiogram import BaseMiddleware
from aiogram.exceptions import TelegramBadRequest
from aiogram.types import CallbackQuery, Message, TelegramObject

from app.keyboards.user import subscription_keyboard
from app.texts.user import subscription_failed_text


class SubscriptionGuardMiddleware(BaseMiddleware):
    async def __call__(
        self,
        handler: Callable[[TelegramObject, dict[str, Any]], Awaitable[Any]],
        event: TelegramObject,
        data: dict[str, Any],
    ) -> Any:
        users_repo = data.get('users_repo')
        subscription_service = data.get('subscription_service')
        bot = data.get('bot')
        admin_ids = set(data.get('admin_ids') or set())

        if not users_repo or not subscription_service or not bot:
            return await handler(event, data)

        user_id: int | None = None

        if isinstance(event, Message):
            if event.chat.type != 'private':
                return await handler(event, data)

            if event.text and (event.text.startswith('/start') or event.text.startswith('/admin')):
                return await handler(event, data)

            user_id = event.from_user.id

        elif isinstance(event, CallbackQuery):
            if not event.message or event.message.chat.type != 'private':
                return await handler(event, data)

            if isinstance(event.data, str) and event.data.startswith('sub:'):
                return await handler(event, data)

            user_id = event.from_user.id

        else:
            return await handler(event, data)

        if user_id in admin_ids:
            return await handler(event, data)

        has_channels = await subscription_service.has_active_channels()
        if not has_channels:
            return await handler(event, data)

        is_subscribed, unsubscribed_channels = await subscription_service.check_user_subscriptions(
            bot=bot,
            user_id=user_id,
        )

        if is_subscribed:
            user = await users_repo.get_by_telegram_id(user_id)
            if user and not user.get('subscription_verified'):
                await users_repo.set_subscription_verified(user_id, True)
            return await handler(event, data)

        await users_repo.set_subscription_verified(user_id, False)

        all_channels = await subscription_service.get_active_channels()
        text = subscription_failed_text(unsubscribed_channels)

        if isinstance(event, Message):
            await event.answer(
                text=text,
                reply_markup=subscription_keyboard(all_channels),
            )
            return None

        if isinstance(event, CallbackQuery):
            try:
                await event.message.edit_text(
                    text=text,
                    reply_markup=subscription_keyboard(all_channels),
                )
            except TelegramBadRequest as e:
                if 'message is not modified' not in str(e):
                    await event.message.answer(
                        text=text,
                        reply_markup=subscription_keyboard(all_channels),
                    )
            except Exception:
                await event.message.answer(
                    text=text,
                    reply_markup=subscription_keyboard(all_channels),
                )

            await event.answer('Avval majburiy obunani tasdiqlang.', show_alert=True)
            return None

        return await handler(event, data)
