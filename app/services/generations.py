from __future__ import annotations

from app.repositories.users import UsersRepository


class GenerationAccessService:
    @staticmethod
    def available_generations(user: dict | None) -> int:
        if not user:
            return 0
        if user.get('bot_access_blocked') or user.get('generation_access_blocked'):
            return 0
        if user.get('generation_unlimited'):
            return 10**9

        first_free = 0 if user.get('free_generation_used', False) else 1
        credits = int(user.get('referral_credits', 0) or 0)
        bonus = int(user.get('bonus_generation_credits', 0) or 0)
        return first_free + credits + bonus

    def has_available_generation(self, user: dict | None) -> bool:
        return self.available_generations(user) > 0

    async def consume_generation(self, users_repo: UsersRepository, telegram_id: int) -> str | None:
        return await users_repo.consume_generation_credit(telegram_id)

    async def restore_consumed_generation(self, users_repo: UsersRepository, telegram_id: int, consumed_from: str) -> None:
        await users_repo.restore_generation_credit(telegram_id, consumed_from)
