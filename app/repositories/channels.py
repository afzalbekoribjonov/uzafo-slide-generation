from __future__ import annotations

from datetime import datetime, timezone
from typing import Any

from motor.motor_asyncio import AsyncIOMotorCollection
from pymongo import ReturnDocument


def utcnow() -> datetime:
    return datetime.now(timezone.utc)


class ChannelsRepository:
    def __init__(self, collection: AsyncIOMotorCollection) -> None:
        self.collection = collection

    async def list_active(self) -> list[dict[str, Any]]:
        cursor = self.collection.find({'is_active': True}).sort('created_at', 1)
        return await cursor.to_list(length=100)

    async def has_active_channels(self) -> bool:
        channel = await self.collection.find_one({'is_active': True})
        return channel is not None

    async def list_all(self, limit: int = 100) -> list[dict[str, Any]]:
        cursor = self.collection.find({}).sort([('is_active', -1), ('created_at', 1)]).limit(limit)
        return await cursor.to_list(length=limit)

    async def get_by_chat_id(self, chat_id: int) -> dict[str, Any] | None:
        return await self.collection.find_one({'chat_id': chat_id})

    async def upsert_channel(
        self,
        *,
        chat_id: int,
        title: str,
        username: str | None,
        invite_link: str | None,
        is_active: bool = True,
    ) -> dict[str, Any]:
        now = utcnow()
        return await self.collection.find_one_and_update(
            {'chat_id': chat_id},
            {
                '$set': {
                    'title': title,
                    'username': username,
                    'invite_link': invite_link,
                    'is_active': bool(is_active),
                    'updated_at': now,
                },
                '$setOnInsert': {
                    'created_at': now,
                },
            },
            upsert=True,
            return_document=ReturnDocument.AFTER,
        )

    async def set_active(self, chat_id: int, value: bool) -> dict[str, Any] | None:
        return await self.collection.find_one_and_update(
            {'chat_id': chat_id},
            {
                '$set': {
                    'is_active': bool(value),
                    'updated_at': utcnow(),
                }
            },
            return_document=ReturnDocument.AFTER,
        )

    async def delete_by_chat_id(self, chat_id: int) -> bool:
        result = await self.collection.delete_one({'chat_id': chat_id})
        return bool(result.deleted_count)
