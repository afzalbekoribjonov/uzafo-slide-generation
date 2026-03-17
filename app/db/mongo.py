from __future__ import annotations

from motor.motor_asyncio import AsyncIOMotorClient, AsyncIOMotorDatabase


class Mongo:
    def __init__(self, uri: str, db_name: str) -> None:
        self._client = AsyncIOMotorClient(uri)
        self.db: AsyncIOMotorDatabase = self._client[db_name]

    async def close(self) -> None:
        self._client.close()
