from __future__ import annotations

from datetime import datetime, timezone
from typing import Any

from motor.motor_asyncio import AsyncIOMotorCollection
from pymongo import ReturnDocument


def utcnow() -> datetime:
    return datetime.now(timezone.utc)


class MagicSettingsRepository:
    SETTINGS_ID = 'magic_slide'
    DEFAULT_PRICE_PER_PRESENTATION = 10000

    def __init__(self, collection: AsyncIOMotorCollection) -> None:
        self.collection = collection

    async def get_settings(self) -> dict[str, Any]:
        now = utcnow()
        return await self.collection.find_one_and_update(
            {'_id': self.SETTINGS_ID},
            {
                '$setOnInsert': {
                    'price_per_presentation': self.DEFAULT_PRICE_PER_PRESENTATION,
                    'maintenance_enabled': False,
                    'created_at': now,
                },
                '$set': {
                    'updated_at': now,
                },
            },
            upsert=True,
            return_document=ReturnDocument.AFTER,
        )

    async def set_price(self, amount_uzs: int) -> dict[str, Any]:
        now = utcnow()
        return await self.collection.find_one_and_update(
            {'_id': self.SETTINGS_ID},
            {
                '$set': {
                    'price_per_presentation': int(amount_uzs),
                    'updated_at': now,
                },
                '$setOnInsert': {
                    'maintenance_enabled': False,
                    'created_at': now,
                },
            },
            upsert=True,
            return_document=ReturnDocument.AFTER,
        )

    async def set_maintenance(self, value: bool) -> dict[str, Any]:
        now = utcnow()
        return await self.collection.find_one_and_update(
            {'_id': self.SETTINGS_ID},
            {
                '$set': {
                    'maintenance_enabled': bool(value),
                    'updated_at': now,
                },
                '$setOnInsert': {
                    'price_per_presentation': self.DEFAULT_PRICE_PER_PRESENTATION,
                    'created_at': now,
                },
            },
            upsert=True,
            return_document=ReturnDocument.AFTER,
        )
