from __future__ import annotations

from app.repositories.referrals import ReferralsRepository
from app.repositories.users import UsersRepository


class ReferralService:
    def __init__(self, referrals_repo: ReferralsRepository, users_repo: UsersRepository) -> None:
        self.referrals_repo = referrals_repo
        self.users_repo = users_repo

    async def register_start_if_valid(self, *, inviter_id: int | None, invitee_id: int, is_new_user: bool) -> None:
        if not inviter_id or not is_new_user:
            return
        if inviter_id == invitee_id:
            return
        inviter = await self.users_repo.get_by_telegram_id(inviter_id)
        if inviter is None:
            return
        existing = await self.referrals_repo.get_by_invitee_id(invitee_id)
        if existing is not None:
            return
        await self.referrals_repo.create_pending(inviter_id=inviter_id, invitee_id=invitee_id)

    async def approve_after_subscription(self, invitee_id: int) -> dict | None:
        referral = await self.referrals_repo.mark_verified_and_counted(invitee_id)
        if referral is None:
            return None

        inviter_id = referral['inviter_id']
        await self.users_repo.increment_referral_reward(inviter_id)

        invitee = await self.users_repo.get_by_telegram_id(invitee_id)
        inviter = await self.users_repo.get_by_telegram_id(inviter_id)

        return {
            'inviter_id': inviter_id,
            'invitee_id': invitee_id,
            'invitee_name': invitee.get('full_name', str(invitee_id)) if invitee else str(invitee_id),
            'referral_credits': inviter.get('referral_credits', 0) if inviter else 0,
            'referral_count': inviter.get('referral_count', 0) if inviter else 0,
        }

    async def list_invited_users(self, inviter_id: int) -> list[dict]:
        return await self.referrals_repo.list_by_inviter(inviter_id)
