from __future__ import annotations

from aiogram.filters import BaseFilter
from aiogram.types import CallbackQuery, Message, TelegramObject


class AdminFilter(BaseFilter):
    async def __call__(self, event: TelegramObject, admin_ids: set[int] | None = None) -> bool:
        if not admin_ids:
            return False

        user_id: int | None = None
        if isinstance(event, Message) and event.from_user:
            user_id = event.from_user.id
        elif isinstance(event, CallbackQuery) and event.from_user:
            user_id = event.from_user.id

        return user_id in set(admin_ids or set())
