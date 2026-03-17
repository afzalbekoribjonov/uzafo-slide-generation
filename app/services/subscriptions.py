from __future__ import annotations

from aiogram import Bot


class SubscriptionService:
    def __init__(self, channels_repo) -> None:
        self.channels_repo = channels_repo

    async def get_active_channels(self) -> list[dict]:
        return await self.channels_repo.list_active()

    async def has_active_channels(self) -> bool:
        return await self.channels_repo.has_active_channels()

    async def check_user_subscriptions(self, bot: Bot, user_id: int) -> tuple[bool, list[dict]]:
        channels = await self.channels_repo.list_active()
        if not channels:
            return True, []

        unsubscribed = []

        for channel in channels:
            chat_id = channel['chat_id']
            try:
                member = await bot.get_chat_member(chat_id=chat_id, user_id=user_id)
                status = member.status
                if status not in ('member', 'administrator', 'creator'):
                    unsubscribed.append(channel)
            except Exception:
                unsubscribed.append(channel)

        return len(unsubscribed) == 0, unsubscribed