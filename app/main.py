from __future__ import annotations

import asyncio
import contextlib
import logging

from aiohttp import web
from aiogram import Bot, Dispatcher
from aiogram.client.default import DefaultBotProperties
from aiogram.enums import ParseMode
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.webhook.aiohttp_server import SimpleRequestHandler, setup_application

from app.config import Settings, get_settings
from app.db.indexes import setup_indexes
from app.db.mongo import Mongo
from app.handlers.admin.panel import router as admin_router
from app.handlers.user.create import router as create_router
from app.handlers.user.menu import router as menu_router
from app.handlers.user.start import router as start_router
from app.handlers.user.subscription import router as subscription_router
from app.middlewares.subscription_guard import SubscriptionGuardMiddleware
from app.middlewares.user_access import UserAccessMiddleware
from app.repositories.channels import ChannelsRepository
from app.repositories.generations import GenerationsRepository
from app.repositories.referrals import ReferralsRepository
from app.repositories.users import UsersRepository
from app.services.admin import AdminService
from app.services.gemini_planner import GeminiPresentationPlanner
from app.services.generation_queue import GenerationQueueService
from app.services.generations import GenerationAccessService
from app.services.pptx_generation import PptxGenerationService
from app.services.referrals import ReferralService
from app.services.subscriptions import SubscriptionService
from app.services.users import UserService

logger = logging.getLogger(__name__)


async def build_runtime(settings: Settings) -> dict:
    bot = Bot(token=settings.bot_token, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
    dp = Dispatcher(storage=MemoryStorage())

    mongo = Mongo(settings.mongodb_uri, settings.mongodb_db)
    await setup_indexes(mongo.db)

    users_repo = UsersRepository(mongo.db.users)
    referrals_repo = ReferralsRepository(mongo.db.referrals)
    channels_repo = ChannelsRepository(mongo.db.mandatory_channels)
    generations_repo = GenerationsRepository(mongo.db.generations)

    admin_ids = set(settings.admins)
    generation_access_service = GenerationAccessService()
    user_service = UserService(users_repo, admin_ids=admin_ids)
    referral_service = ReferralService(referrals_repo, users_repo)
    subscription_service = SubscriptionService(channels_repo)
    gemini_planner = GeminiPresentationPlanner(
        api_key=settings.gemini_api_key,
        model_name=settings.gemini_model,
        max_retries=settings.gemini_max_retries,
        initial_backoff_seconds=settings.gemini_initial_backoff_seconds,
    )
    pptx_generation_service = PptxGenerationService(gemini_planner=gemini_planner)
    generation_queue_service = GenerationQueueService(
        generations_repo=generations_repo,
        users_repo=users_repo,
        pptx_generation_service=pptx_generation_service,
        poll_interval_seconds=settings.generation_worker_poll_seconds,
        start_cooldown_seconds=settings.generation_start_cooldown_seconds,
    )

    me = await bot.get_me()
    admin_service = AdminService(
        users_repo=users_repo,
        generations_repo=generations_repo,
        generation_access_service=generation_access_service,
        bot_username=me.username or 'your_bot',
    )

    dp['users_repo'] = users_repo
    dp['referrals_repo'] = referrals_repo
    dp['channels_repo'] = channels_repo
    dp['generations_repo'] = generations_repo
    dp['user_service'] = user_service
    dp['referral_service'] = referral_service
    dp['generation_access_service'] = generation_access_service
    dp['generation_queue_service'] = generation_queue_service
    dp['subscription_service'] = subscription_service
    dp['admin_service'] = admin_service
    dp['admin_ids'] = admin_ids
    dp['bot_username'] = me.username or 'your_bot'
    dp['support_contact'] = settings.support_contact

    user_access = UserAccessMiddleware()
    subscription_guard = SubscriptionGuardMiddleware()
    dp.message.middleware(user_access)
    dp.callback_query.middleware(user_access)
    dp.message.middleware(subscription_guard)
    dp.callback_query.middleware(subscription_guard)

    dp.include_router(admin_router)
    dp.include_router(start_router)
    dp.include_router(subscription_router)
    dp.include_router(create_router)
    dp.include_router(menu_router)

    return {
        'settings': settings,
        'bot': bot,
        'dp': dp,
        'mongo': mongo,
        'generation_queue_service': generation_queue_service,
        'bot_username': me.username or 'your_bot',
    }


async def generation_worker_context(app: web.Application):
    bot: Bot = app['bot']
    generation_queue_service: GenerationQueueService = app['generation_queue_service']
    mongo: Mongo = app['mongo']

    worker_task = asyncio.create_task(generation_queue_service.run_worker(bot), name='generation-queue-worker')
    app['generation_worker_task'] = worker_task
    logger.info('Generation queue worker started.')

    try:
        yield
    finally:
        worker_task.cancel()
        with contextlib.suppress(asyncio.CancelledError):
            await worker_task
        logger.info('Generation queue worker stopped.')

        await mongo.close()
        await bot.session.close()


async def healthcheck(_: web.Request) -> web.Response:
    return web.json_response({'status': 'ok'})


async def create_webhook_app() -> web.Application:
    settings = get_settings()
    runtime = await build_runtime(settings)

    bot: Bot = runtime['bot']
    dp: Dispatcher = runtime['dp']

    async def on_startup(bot: Bot) -> None:
        webhook_url = settings.webhook_url
        await bot.set_webhook(
            url=webhook_url,
            secret_token=settings.webhook_secret,
            max_connections=settings.webhook_max_connections,
            drop_pending_updates=settings.webhook_drop_pending_updates,
            allowed_updates=dp.resolve_used_update_types(),
        )
        logger.info('Webhook configured: %s', webhook_url)

    dp.startup.register(on_startup)

    app = web.Application()
    app.update(runtime)
    app.cleanup_ctx.append(generation_worker_context)
    app.router.add_get('/', healthcheck)
    app.router.add_get('/healthz', healthcheck)

    webhook_requests_handler = SimpleRequestHandler(
        dispatcher=dp,
        bot=bot,
        secret_token=settings.webhook_secret,
    )
    webhook_requests_handler.register(app, path=settings.webhook_path)
    setup_application(app, dp, bot=bot)

    logger.info(
        'Starting webhook server on %s:%s with path %s',
        settings.web_server_host,
        settings.web_server_port,
        settings.webhook_path,
    )
    return app


async def run_polling() -> None:
    settings = get_settings()
    runtime = await build_runtime(settings)

    bot: Bot = runtime['bot']
    dp: Dispatcher = runtime['dp']
    mongo: Mongo = runtime['mongo']
    generation_queue_service: GenerationQueueService = runtime['generation_queue_service']

    worker_task = asyncio.create_task(generation_queue_service.run_worker(bot), name='generation-queue-worker')
    logger.info('Starting polling mode.')

    try:
        await dp.start_polling(bot, allowed_updates=dp.resolve_used_update_types())
    finally:
        worker_task.cancel()
        with contextlib.suppress(asyncio.CancelledError):
            await worker_task
        await mongo.close()
        await bot.session.close()


def main() -> None:
    logging.basicConfig(level=logging.WARNING, format='%(asctime)s | %(levelname)s | %(name)s | %(message)s')

    settings = get_settings()
    if settings.app_mode == 'webhook':
        web.run_app(
            create_webhook_app(),
            host=settings.web_server_host,
            port=settings.web_server_port,
        )
        return

    asyncio.run(run_polling())


if __name__ == '__main__':
    main()
