from __future__ import annotations

from typing import Any

from motor.motor_asyncio import AsyncIOMotorCollection


class ChannelsRepository:
    def __init__(self, collection: AsyncIOMotorCollection) -> None:
        self.collection = collection

    async def list_active(self) -> list[dict[str, Any]]:
        cursor = self.collection.find({'is_active': True}).sort('created_at', 1)
        return await cursor.to_list(length=100)

    async def has_active_channels(self) -> bool:
        channel = await self.collection.find_one({'is_active': True})
        return channel is not None