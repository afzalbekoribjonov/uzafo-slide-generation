from __future__ import annotations

from datetime import datetime, timezone
from typing import Any

from bson import ObjectId
from motor.motor_asyncio import AsyncIOMotorCollection
from pymongo import DESCENDING, ReturnDocument


def utcnow() -> datetime:
    return datetime.now(timezone.utc)


class MagicTopupsRepository:
    def __init__(self, collection: AsyncIOMotorCollection) -> None:
        self.collection = collection

    async def create_pending(
        self,
        *,
        telegram_id: int,
        full_name: str,
        username: str | None,
        amount_uzs: int,
        receipt_type: str,
        receipt_file_id: str,
        receipt_file_unique_id: str | None,
        receipt_file_name: str | None,
        receipt_mime_type: str | None,
        receipt_caption: str | None,
        cards_snapshot: list[dict[str, Any]],
    ) -> dict[str, Any]:
        now = utcnow()
        document = {
            'telegram_id': telegram_id,
            'full_name': full_name,
            'username': username,
            'amount_uzs': int(amount_uzs),
            'status': 'pending',
            'receipt_type': receipt_type,
            'receipt_file_id': receipt_file_id,
            'receipt_file_unique_id': receipt_file_unique_id,
            'receipt_file_name': receipt_file_name,
            'receipt_mime_type': receipt_mime_type,
            'receipt_caption': receipt_caption,
            'cards_snapshot': cards_snapshot,
            'admin_notifications': [],
            'reviewed_by': None,
            'reviewed_at': None,
            'review_decision': None,
            'created_at': now,
            'updated_at': now,
        }
        result = await self.collection.insert_one(document)
        document['_id'] = result.inserted_id
        return document

    async def get_by_id(self, topup_id: str) -> dict[str, Any] | None:
        return await self.collection.find_one({'_id': ObjectId(topup_id)})

    async def list_pending(self, limit: int = 20) -> list[dict[str, Any]]:
        cursor = self.collection.find({'status': 'pending'}).sort([('created_at', DESCENDING)]).limit(limit)
        return await cursor.to_list(length=limit)

    async def count_pending(self) -> int:
        return await self.collection.count_documents({'status': 'pending'})

    async def add_admin_notification(self, topup_id: str, *, chat_id: int, message_id: int) -> None:
        await self.collection.update_one(
            {'_id': ObjectId(topup_id)},
            {
                '$push': {
                    'admin_notifications': {
                        'chat_id': int(chat_id),
                        'message_id': int(message_id),
                    }
                },
                '$set': {
                    'updated_at': utcnow(),
                },
            },
        )

    async def mark_approved(self, topup_id: str, *, admin_id: int, admin_name: str) -> dict[str, Any] | None:
        now = utcnow()
        return await self.collection.find_one_and_update(
            {'_id': ObjectId(topup_id), 'status': 'pending'},
            {
                '$set': {
                    'status': 'approved',
                    'reviewed_by': {
                        'telegram_id': int(admin_id),
                        'full_name': str(admin_name),
                    },
                    'reviewed_at': now,
                    'review_decision': 'approved',
                    'updated_at': now,
                }
            },
            return_document=ReturnDocument.AFTER,
        )

    async def mark_rejected(self, topup_id: str, *, admin_id: int, admin_name: str) -> dict[str, Any] | None:
        now = utcnow()
        return await self.collection.find_one_and_update(
            {'_id': ObjectId(topup_id), 'status': 'pending'},
            {
                '$set': {
                    'status': 'rejected',
                    'reviewed_by': {
                        'telegram_id': int(admin_id),
                        'full_name': str(admin_name),
                    },
                    'reviewed_at': now,
                    'review_decision': 'rejected',
                    'updated_at': now,
                }
            },
            return_document=ReturnDocument.AFTER,
        )
