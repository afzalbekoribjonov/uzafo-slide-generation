from __future__ import annotations

from datetime import datetime, timezone
from typing import Any

from bson import ObjectId
from motor.motor_asyncio import AsyncIOMotorCollection
from pymongo import DESCENDING, ReturnDocument


STATUSES_ACTIVE = ('queued', 'processing')


def utcnow() -> datetime:
    return datetime.now(timezone.utc)


class GenerationsRepository:
    def __init__(self, collection: AsyncIOMotorCollection) -> None:
        self.collection = collection

    async def create_job(
        self,
        *,
        telegram_id: int,
        full_name: str,
        username: str | None,
        payload: dict[str, Any],
        consumed_from: str,
        status_chat_id: int | None = None,
        status_message_id: int | None = None,
    ) -> dict[str, Any]:
        document = {
            'telegram_id': telegram_id,
            'full_name': full_name,
            'username': username,
            'payload': payload,
            'consumed_from': consumed_from,
            'status': 'queued',
            'status_chat_id': status_chat_id,
            'status_message_id': status_message_id,
            'created_at': utcnow(),
            'started_at': None,
            'finished_at': None,
            'result_file_path': None,
            'error': None,
            'attempts': 0,
        }
        result = await self.collection.insert_one(document)
        document['_id'] = result.inserted_id
        return document

    async def get_by_id(self, generation_id: ObjectId | str) -> dict[str, Any] | None:
        object_id = generation_id if isinstance(generation_id, ObjectId) else ObjectId(generation_id)
        return await self.collection.find_one({'_id': object_id})

    async def get_active_job_for_user(self, telegram_id: int) -> dict[str, Any] | None:
        return await self.collection.find_one(
            {
                'telegram_id': telegram_id,
                'status': {'$in': list(STATUSES_ACTIVE)},
            },
            sort=[('created_at', 1)],
        )

    async def count_ahead_in_queue(self, generation_id: ObjectId | str) -> int:
        job = await self.get_by_id(generation_id)
        if not job:
            return 0

        queued_ahead = await self.collection.count_documents(
            {
                'status': 'queued',
                'created_at': {'$lt': job['created_at']},
            }
        )
        processing_count = await self.collection.count_documents({'status': 'processing'})
        if job.get('status') == 'processing' and processing_count > 0:
            processing_count -= 1
        return queued_ahead + processing_count

    async def claim_next_queued_job(self) -> dict[str, Any] | None:
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

    async def set_status_message(self, generation_id: ObjectId | str, *, chat_id: int, message_id: int) -> None:
        object_id = generation_id if isinstance(generation_id, ObjectId) else ObjectId(generation_id)
        await self.collection.update_one(
            {'_id': object_id},
            {
                '$set': {
                    'status_chat_id': chat_id,
                    'status_message_id': message_id,
                }
            },
        )

    async def mark_done(self, generation_id: ObjectId | str, *, result_file_path: str | None = None) -> None:
        object_id = generation_id if isinstance(generation_id, ObjectId) else ObjectId(generation_id)
        await self.collection.update_one(
            {'_id': object_id},
            {
                '$set': {
                    'status': 'done',
                    'finished_at': utcnow(),
                    'result_file_path': result_file_path,
                }
            },
        )

    async def mark_failed(self, generation_id: ObjectId | str, *, error: str) -> None:
        object_id = generation_id if isinstance(generation_id, ObjectId) else ObjectId(generation_id)
        await self.collection.update_one(
            {'_id': object_id},
            {
                '$set': {
                    'status': 'failed',
                    'finished_at': utcnow(),
                    'error': error,
                }
            },
        )

    async def requeue_processing_jobs(self) -> int:
        result = await self.collection.update_many(
            {'status': 'processing'},
            {
                '$set': {
                    'status': 'queued',
                    'started_at': None,
                    'error': 'Bot qayta ishga tushgani uchun so‘rov navbatga qaytarildi.',
                }
            },
        )
        return int(result.modified_count)

    async def count_matching(self, query: dict[str, Any] | None = None) -> int:
        return await self.collection.count_documents(query or {})

    async def get_status_counts(self) -> dict[str, int]:
        pipeline = [
            {
                '$group': {
                    '_id': '$status',
                    'count': {'$sum': 1},
                }
            }
        ]
        items = await self.collection.aggregate(pipeline).to_list(length=20)
        result = {'queued': 0, 'processing': 0, 'done': 0, 'failed': 0}
        for item in items:
            status = str(item.get('_id') or '')
            result[status] = int(item.get('count', 0) or 0)
        return result

    async def list_recent_failures(self, limit: int = 5) -> list[dict[str, Any]]:
        cursor = (
            self.collection.find({'status': 'failed'})
            .sort([('finished_at', DESCENDING), ('created_at', DESCENDING)])
            .limit(limit)
        )
        return await cursor.to_list(length=limit)
