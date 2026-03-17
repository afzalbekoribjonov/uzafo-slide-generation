from __future__ import annotations

from aiogram.types import User as TelegramUser

from app.repositories.users import UsersRepository


class UserService:
    def __init__(self, users_repo: UsersRepository, *, admin_ids: set[int] | None = None) -> None:
        self.users_repo = users_repo
        self.admin_ids = set(admin_ids or set())

    def is_admin_id(self, telegram_id: int) -> bool:
        return telegram_id in self.admin_ids

    async def get_or_create_user(self, tg_user: TelegramUser, invited_by: int | None = None) -> tuple[dict, bool]:
        is_admin = self.is_admin_id(tg_user.id)
        user = await self.users_repo.get_by_telegram_id(tg_user.id)
        if user:
            await self.users_repo.touch(
                tg_user.id,
                full_name=tg_user.full_name,
                username=tg_user.username,
                is_admin=is_admin,
            )
            fresh_user = await self.users_repo.get_by_telegram_id(tg_user.id)
            return fresh_user, False

        created = await self.users_repo.create(
            telegram_id=tg_user.id,
            full_name=tg_user.full_name,
            username=tg_user.username,
            invited_by=invited_by,
            is_admin=is_admin,
        )
        return created, True

    async def get_user(self, telegram_id: int) -> dict | None:
        user = await self.users_repo.get_by_telegram_id(telegram_id)
        if user and user.get('is_admin') != self.is_admin_id(telegram_id):
            await self.users_repo.sync_admin_flag(telegram_id, self.is_admin_id(telegram_id))
            user = await self.users_repo.get_by_telegram_id(telegram_id)
        return user

    async def mark_subscription_verified(self, telegram_id: int) -> None:
        await self.users_repo.mark_subscription_verified(telegram_id)

    async def set_subscription_verified(self, telegram_id: int, value: bool) -> None:
        await self.users_repo.set_subscription_verified(telegram_id, value)
