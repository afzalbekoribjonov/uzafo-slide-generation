from __future__ import annotations

from datetime import datetime, timezone
from typing import Any

from motor.motor_asyncio import AsyncIOMotorCollection
from pymongo import ReturnDocument


def utcnow() -> datetime:
    return datetime.now(timezone.utc)


class ReferralsRepository:
    def __init__(self, collection: AsyncIOMotorCollection) -> None:
        self.collection = collection

    async def get_by_invitee_id(self, invitee_id: int) -> dict[str, Any] | None:
        return await self.collection.find_one({'invitee_id': invitee_id})

    async def create_pending(self, *, inviter_id: int, invitee_id: int) -> None:
        await self.collection.insert_one(
            {
                'inviter_id': inviter_id,
                'invitee_id': invitee_id,
                'start_detected': True,
                'subscription_verified': False,
                'counted': False,
                'created_at': utcnow(),
                'counted_at': None,
            }
        )

    async def mark_verified_and_counted(self, invitee_id: int) -> dict[str, Any] | None:
        referral = await self.collection.find_one_and_update(
            {'invitee_id': invitee_id, 'counted': False},
            {
                '$set': {
                    'subscription_verified': True,
                    'counted': True,
                    'counted_at': utcnow(),
                }
            },
            return_document=ReturnDocument.AFTER,
        )
        return referral

    async def list_by_inviter(self, inviter_id: int, limit: int = 20) -> list[dict[str, Any]]:
        cursor = self.collection.find({'inviter_id': inviter_id}).sort('created_at', -1).limit(limit)
        return await cursor.to_list(length=limit)
