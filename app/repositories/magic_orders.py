from __future__ import annotations

from datetime import datetime, timezone
from typing import Any

from bson import ObjectId
from motor.motor_asyncio import AsyncIOMotorCollection
from pymongo import DESCENDING, ReturnDocument


STATUSES_ACTIVE = ('queued', 'processing')


def utcnow() -> datetime:
    return datetime.now(timezone.utc)


class MagicOrdersRepository:
    def __init__(self, collection: AsyncIOMotorCollection) -> None:
        self.collection = collection

    async def create_job(
        self,
        *,
        telegram_id: int,
        full_name: str,
        username: str | None,
        payload: dict[str, Any],
        template_id: str,
        template_name: str,
        category: str,
        output_slide_target: int,
        price_uzs_snapshot: int,
        status_chat_id: int | None = None,
        status_message_id: int | None = None,
    ) -> dict[str, Any]:
        document = {
            'telegram_id': telegram_id,
            'full_name': full_name,
            'username': username,
            'payload': payload,
            'template_id': template_id,
            'template_name': template_name,
            'category': category,
            'output_slide_target': int(output_slide_target),
            'price_uzs_snapshot': int(price_uzs_snapshot),
            'status': 'queued',
            'status_chat_id': status_chat_id,
            'status_message_id': status_message_id,
            'created_at': utcnow(),
            'started_at': None,
            'finished_at': None,
            'result_file_path': None,
            'error': None,
            'attempts': 0,
            'charged_amount_uzs': 0,
        }
        result = await self.collection.insert_one(document)
        document['_id'] = result.inserted_id
        return document

    async def get_by_id(self, order_id: ObjectId | str) -> dict[str, Any] | None:
        object_id = order_id if isinstance(order_id, ObjectId) else ObjectId(order_id)
        return await self.collection.find_one({'_id': object_id})

    async def get_active_order_for_user(self, telegram_id: int) -> dict[str, Any] | None:
        return await self.collection.find_one(
            {
                'telegram_id': telegram_id,
                'status': {'$in': list(STATUSES_ACTIVE)},
            },
            sort=[('created_at', 1)],
        )

    async def count_ahead_in_queue(self, order_id: ObjectId | str) -> int:
        order = await self.get_by_id(order_id)
        if not order:
            return 0

        queued_ahead = await self.collection.count_documents(
            {
                'status': 'queued',
                'created_at': {'$lt': order['created_at']},
            }
        )
        processing_count = await self.collection.count_documents({'status': 'processing'})
        if order.get('status') == 'processing' and processing_count > 0:
            processing_count -= 1
        return queued_ahead + processing_count

    async def claim_next_queued_order(self) -> dict[str, Any] | None:
        return await self.collection.find_one_and_update(
            {'status': 'queued'},
            {
                '$set': {
                    'status': 'processing',
                    'started_at': utcnow(),
                    'error': None,
                },
                '$inc': {
                    'attempts': 1,
                },
            },
            sort=[('created_at', 1)],
            return_document=ReturnDocument.AFTER,
        )

    async def set_status_message(self, order_id: ObjectId | str, *, chat_id: int, message_id: int) -> None:
        object_id = order_id if isinstance(order_id, ObjectId) else ObjectId(order_id)
        await self.collection.update_one(
            {'_id': object_id},
            {
                '$set': {
                    'status_chat_id': chat_id,
                    'status_message_id': message_id,
                    'updated_at': utcnow(),
                }
            },
        )

    async def mark_done(
        self,
        order_id: ObjectId | str,
        *,
        result_file_path: str | None = None,
        charged_amount_uzs: int = 0,
    ) -> None:
        object_id = order_id if isinstance(order_id, ObjectId) else ObjectId(order_id)
        await self.collection.update_one(
            {'_id': object_id},
            {
                '$set': {
                    'status': 'done',
                    'finished_at': utcnow(),
                    'result_file_path': result_file_path,
                    'charged_amount_uzs': int(charged_amount_uzs),
                    'updated_at': utcnow(),
                }
            },
        )

    async def mark_failed(self, order_id: ObjectId | str, *, error: str) -> None:
        object_id = order_id if isinstance(order_id, ObjectId) else ObjectId(order_id)
        await self.collection.update_one(
            {'_id': object_id},
            {
                '$set': {
                    'status': 'failed',
                    'finished_at': utcnow(),
                    'error': error,
                    'updated_at': utcnow(),
                }
            },
        )

    async def requeue_processing_orders(self) -> int:
        result = await self.collection.update_many(
            {'status': 'processing'},
            {
                '$set': {
                    'status': 'queued',
                    'started_at': None,
                    'error': 'Bot qayta ishga tushgani uchun premium buyurtma navbatga qaytarildi.',
                    'updated_at': utcnow(),
                }
            },
        )
        return int(result.modified_count)

    async def list_recent_failures(self, limit: int = 5) -> list[dict[str, Any]]:
        cursor = (
            self.collection.find({'status': 'failed'})
            .sort([('finished_at', DESCENDING), ('created_at', DESCENDING)])
            .limit(limit)
        )
        return await cursor.to_list(length=limit)
