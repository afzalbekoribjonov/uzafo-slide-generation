from __future__ import annotations

from typing import Any, Awaitable, Callable

from aiogram import BaseMiddleware
from aiogram.types import CallbackQuery, Message, TelegramObject

from app.texts.user import bot_access_blocked_alert_text, bot_access_blocked_text


class UserAccessMiddleware(BaseMiddleware):
    async def __call__(
        self,
        handler: Callable[[TelegramObject, dict[str, Any]], Awaitable[Any]],
        event: TelegramObject,
        data: dict[str, Any],
    ) -> Any:
        users_repo = data.get('users_repo')
        admin_ids = set(data.get('admin_ids') or set())

        if not users_repo:
            return await handler(event, data)

        user_id: int | None = None
        chat_type: str | None = None

        if isinstance(event, Message) and event.from_user:
            user_id = event.from_user.id
            chat_type = event.chat.type
        elif isinstance(event, CallbackQuery) and event.from_user and event.message:
            user_id = event.from_user.id
            chat_type = event.message.chat.type

        if not user_id or chat_type != 'private' or user_id in admin_ids:
            return await handler(event, data)

        user = await users_repo.get_by_telegram_id(user_id)
        if not user or not user.get('bot_access_blocked'):
            return await handler(event, data)

        if isinstance(event, Message):
            await event.answer(bot_access_blocked_text())
            return None

        if isinstance(event, CallbackQuery):
            await event.answer(bot_access_blocked_alert_text(), show_alert=True)
            return None

        return await handler(event, data)
