from __future__ import annotations

from datetime import datetime, timezone
from typing import Any

from motor.motor_asyncio import AsyncIOMotorCollection
from pymongo import ReturnDocument


def utcnow() -> datetime:
    return datetime.now(timezone.utc)


class MagicAccountsRepository:
    def __init__(self, collection: AsyncIOMotorCollection) -> None:
        self.collection = collection

    async def ensure_account(self, telegram_id: int) -> dict[str, Any]:
        now = utcnow()
        return await self.collection.find_one_and_update(
            {'telegram_id': telegram_id},
            {
                '$setOnInsert': {
                    'telegram_id': telegram_id,
                    'balance_uzs': 0,
                    'total_topped_up_uzs': 0,
                    'total_spent_uzs': 0,
                    'presentations_paid': 0,
                    'created_at': now,
                },
                '$set': {
                    'updated_at': now,
                },
            },
            upsert=True,
            return_document=ReturnDocument.AFTER,
        )

    async def get_by_telegram_id(self, telegram_id: int) -> dict[str, Any] | None:
        return await self.collection.find_one({'telegram_id': telegram_id})

    async def add_balance(self, telegram_id: int, amount_uzs: int) -> dict[str, Any]:
        now = utcnow()
        return await self.collection.find_one_and_update(
            {'telegram_id': telegram_id},
            {
                '$inc': {
                    'balance_uzs': int(amount_uzs),
                    'total_topped_up_uzs': int(amount_uzs),
                },
                '$set': {
                    'updated_at': now,
                    'last_topup_at': now,
                },
                '$setOnInsert': {
                    'created_at': now,
                    'total_spent_uzs': 0,
                    'presentations_paid': 0,
                },
            },
            upsert=True,
            return_document=ReturnDocument.AFTER,
        )

    async def spend_balance(self, telegram_id: int, amount_uzs: int) -> dict[str, Any] | None:
        now = utcnow()
        return await self.collection.find_one_and_update(
            {
                'telegram_id': telegram_id,
                'balance_uzs': {'$gte': int(amount_uzs)},
            },
            {
                '$inc': {
                    'balance_uzs': -int(amount_uzs),
                    'total_spent_uzs': int(amount_uzs),
                    'presentations_paid': 1,
                },
                '$set': {
                    'updated_at': now,
                    'last_spent_at': now,
                },
            },
            return_document=ReturnDocument.AFTER,
        )
