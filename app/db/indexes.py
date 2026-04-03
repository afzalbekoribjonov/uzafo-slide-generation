from __future__ import annotations

from motor.motor_asyncio import AsyncIOMotorDatabase
from pymongo import ASCENDING, DESCENDING, IndexModel


async def setup_indexes(db: AsyncIOMotorDatabase) -> None:
    await db.users.create_indexes(
        [
            IndexModel([('telegram_id', ASCENDING)], unique=True),
            IndexModel([('username', ASCENDING)]),
            IndexModel([('created_at', ASCENDING)]),
            IndexModel([('last_active_at', DESCENDING)]),
            IndexModel([('subscription_verified', ASCENDING)]),
            IndexModel([('referral_count', DESCENDING)]),
            IndexModel([('successful_generations', DESCENDING)]),
            IndexModel([('generation_access_blocked', ASCENDING)]),
            IndexModel([('bot_access_blocked', ASCENDING)]),
        ]
    )

    await db.referrals.create_indexes(
        [
            IndexModel([('inviter_id', ASCENDING)]),
            IndexModel([('invitee_id', ASCENDING)], unique=True),
            IndexModel([('counted', ASCENDING)]),
            IndexModel([('created_at', ASCENDING)]),
        ]
    )

    await db.mandatory_channels.create_indexes(
        [
            IndexModel([('chat_id', ASCENDING)], unique=True),
            IndexModel([('username', ASCENDING)]),
            IndexModel([('is_active', ASCENDING)]),
        ]
    )

    await db.generations.create_indexes(
        [
            IndexModel([('telegram_id', ASCENDING)]),
            IndexModel([('status', ASCENDING), ('created_at', ASCENDING)]),
            IndexModel([('created_at', ASCENDING)]),
            IndexModel([('finished_at', DESCENDING)]),
        ]
    )

    await db.magic_accounts.create_indexes(
        [
            IndexModel([('telegram_id', ASCENDING)], unique=True),
            IndexModel([('balance_uzs', DESCENDING)]),
            IndexModel([('updated_at', DESCENDING)]),
        ]
    )

    await db.magic_cards.create_indexes(
        [
            IndexModel([('is_active', ASCENDING)]),
            IndexModel([('created_at', ASCENDING)]),
        ]
    )

    await db.magic_topups.create_indexes(
        [
            IndexModel([('telegram_id', ASCENDING)]),
            IndexModel([('status', ASCENDING), ('created_at', DESCENDING)]),
            IndexModel([('created_at', DESCENDING)]),
        ]
    )

    await db.magic_orders.create_indexes(
        [
            IndexModel([('telegram_id', ASCENDING)]),
            IndexModel([('status', ASCENDING), ('created_at', DESCENDING)]),
            IndexModel([('created_at', DESCENDING)]),
        ]
    )
