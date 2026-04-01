from __future__ import annotations

import logging
from typing import Any, Awaitable, Callable

from aiogram import BaseMiddleware
from aiogram.types import CallbackQuery, Message, TelegramObject
from pymongo.errors import PyMongoError

from app.texts.user import technical_maintenance_alert_text, technical_maintenance_text

logger = logging.getLogger(__name__)


class DatabaseResilienceMiddleware(BaseMiddleware):
    async def __call__(
        self,
        handler: Callable[[TelegramObject, dict[str, Any]], Awaitable[Any]],
        event: TelegramObject,
        data: dict[str, Any],
    ) -> Any:
        try:
            return await handler(event, data)
        except (PyMongoError, TimeoutError) as exc:
            logger.warning('Database operation failed during update handling: %s', exc)
            await self._notify_user(event)
            return None

    @staticmethod
    async def _notify_user(event: TelegramObject) -> None:
        try:
            if isinstance(event, Message):
                await event.answer(technical_maintenance_text())
                return

            if isinstance(event, CallbackQuery):
                if event.message:
                    await event.message.answer(technical_maintenance_text())
                await event.answer(technical_maintenance_alert_text(), show_alert=True)
        except Exception:
            logger.exception('Failed to send technical-maintenance notice to the user.')
