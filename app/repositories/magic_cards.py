from __future__ import annotations

from datetime import datetime, timezone
from typing import Any

from bson import ObjectId
from motor.motor_asyncio import AsyncIOMotorCollection
from pymongo import ReturnDocument


def utcnow() -> datetime:
    return datetime.now(timezone.utc)


class MagicCardsRepository:
    def __init__(self, collection: AsyncIOMotorCollection) -> None:
        self.collection = collection

    async def list_all(self, limit: int = 50) -> list[dict[str, Any]]:
        cursor = self.collection.find({}).sort([('is_active', -1), ('created_at', 1)]).limit(limit)
        return await cursor.to_list(length=limit)

    async def list_active(self, limit: int = 20) -> list[dict[str, Any]]:
        cursor = self.collection.find({'is_active': True}).sort('created_at', 1).limit(limit)
        return await cursor.to_list(length=limit)

    async def get_by_id(self, card_id: str) -> dict[str, Any] | None:
        return await self.collection.find_one({'_id': ObjectId(card_id)})

    async def create_card(self, *, card_number: str, card_holder: str, is_active: bool = True) -> dict[str, Any]:
        document = {
            'card_number': str(card_number),
            'card_holder': str(card_holder),
            'is_active': bool(is_active),
            'created_at': utcnow(),
            'updated_at': utcnow(),
        }
        result = await self.collection.insert_one(document)
        document['_id'] = result.inserted_id
        return document

    async def set_active(self, card_id: str, value: bool) -> dict[str, Any] | None:
        return await self.collection.find_one_and_update(
            {'_id': ObjectId(card_id)},
            {
                '$set': {
                    'is_active': bool(value),
                    'updated_at': utcnow(),
                }
            },
            return_document=ReturnDocument.AFTER,
        )

    async def delete_by_id(self, card_id: str) -> bool:
        result = await self.collection.delete_one({'_id': ObjectId(card_id)})
        return bool(result.deleted_count)
