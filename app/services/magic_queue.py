from __future__ import annotations

import asyncio
import logging
from pathlib import Path

from aiogram import Bot
from aiogram.exceptions import TelegramBadRequest
from aiogram.types import FSInputFile

from app.keyboards.magic import magic_account_keyboard
from app.repositories.magic_accounts import MagicAccountsRepository
from app.repositories.magic_orders import MagicOrdersRepository
from app.repositories.users import UsersRepository
from app.services.magic_generation import MagicGenerationService
from app.texts.magic import (
    magic_order_balance_missing_text,
    magic_order_charge_issue_text,
    magic_order_done_text,
    magic_order_failed_text,
    magic_order_progress_text,
    magic_order_success_caption,
)

logger = logging.getLogger(__name__)


class MagicOrderQueueService:
    def __init__(
        self,
        *,
        orders_repo: MagicOrdersRepository,
        accounts_repo: MagicAccountsRepository,
        users_repo: UsersRepository,
        generation_service: MagicGenerationService,
        poll_interval_seconds: int = 3,
        start_cooldown_seconds: int = 65,
    ) -> None:
        self.orders_repo = orders_repo
        self.accounts_repo = accounts_repo
        self.users_repo = users_repo
        self.generation_service = generation_service
        self.poll_interval_seconds = max(1, int(poll_interval_seconds))
        self.start_cooldown_seconds = max(0, int(start_cooldown_seconds))
        self._last_started_monotonic: float | None = None

    async def run_worker(self, bot: Bot) -> None:
        while True:
            try:
                requeued = await self.orders_repo.requeue_processing_orders()
                if requeued:
                    logger.warning('Requeued %s premium processing orders after restart.', requeued)
                break
            except asyncio.CancelledError:
                raise
            except Exception:
                logger.exception('Magic order worker could not initialize queue state. Retrying...')
                await asyncio.sleep(self.poll_interval_seconds)

        while True:
            try:
                await self._respect_cooldown()
                order = await self.orders_repo.claim_next_queued_order()
                if not order:
                    await asyncio.sleep(self.poll_interval_seconds)
                    continue

                self._last_started_monotonic = asyncio.get_running_loop().time()
                await self._process_order(bot, order)
            except asyncio.CancelledError:
                raise
            except Exception:
                logger.exception('Magic order worker loop failed unexpectedly.')
                await asyncio.sleep(self.poll_interval_seconds)

    async def _respect_cooldown(self) -> None:
        if self._last_started_monotonic is None or self.start_cooldown_seconds <= 0:
            return

        elapsed = asyncio.get_running_loop().time() - self._last_started_monotonic
        remaining = self.start_cooldown_seconds - elapsed
        if remaining > 0:
            await asyncio.sleep(remaining)

    async def _process_order(self, bot: Bot, order: dict) -> None:
        order_id = order['_id']
        telegram_id = int(order['telegram_id'])
        payload = dict(order.get('payload') or {})
        price_uzs = int(order.get('price_uzs_snapshot', 0) or 0)
        status_chat_id = order.get('status_chat_id')
        status_message_id = order.get('status_message_id')
        file_path: str | None = None

        try:
            account = await self.accounts_repo.ensure_account(telegram_id)
            if int(account.get('balance_uzs', 0) or 0) < price_uzs:
                raise RuntimeError('BALANCE_TOO_LOW')

            status_chat_id, status_message_id = await self._push_progress(
                bot=bot,
                order_id=order_id,
                fallback_chat_id=telegram_id,
                payload=payload,
                percent=14,
                stage_key='queued',
                status_chat_id=status_chat_id,
                status_message_id=status_message_id,
            )
            status_chat_id, status_message_id = await self._push_progress(
                bot=bot,
                order_id=order_id,
                fallback_chat_id=telegram_id,
                payload=payload,
                percent=36,
                stage_key='analysis',
                status_chat_id=status_chat_id,
                status_message_id=status_message_id,
            )
            file_path, template = await asyncio.to_thread(self.generation_service.generate, payload)
            status_chat_id, status_message_id = await self._push_progress(
                bot=bot,
                order_id=order_id,
                fallback_chat_id=telegram_id,
                payload=payload,
                percent=76,
                stage_key='rendering',
                status_chat_id=status_chat_id,
                status_message_id=status_message_id,
            )

            document = FSInputFile(file_path)
            await bot.send_document(
                chat_id=telegram_id,
                document=document,
                caption=magic_order_success_caption(
                    payload=payload,
                    template_name=template.template_name,
                    price_uzs=price_uzs,
                ),
            )

            account_after_charge = await self.accounts_repo.spend_balance(telegram_id, price_uzs)
            await self.orders_repo.mark_done(
                order_id,
                result_file_path=file_path,
                charged_amount_uzs=price_uzs if account_after_charge else 0,
            )
            await self.users_repo.increment_successful_generation(telegram_id)

            await self._push_progress(
                bot=bot,
                order_id=order_id,
                fallback_chat_id=telegram_id,
                payload=payload,
                percent=100,
                stage_key='done',
                status_chat_id=status_chat_id,
                status_message_id=status_message_id,
                override_text=magic_order_done_text(
                    template_name=template.template_name,
                    price_uzs=price_uzs,
                    balance_uzs=(account_after_charge or {}).get('balance_uzs', account.get('balance_uzs', 0)),
                    charged=bool(account_after_charge),
                ),
            )
            if not account_after_charge:
                await bot.send_message(
                    telegram_id,
                    magic_order_charge_issue_text(),
                    reply_markup=magic_account_keyboard(),
                )
            self._cleanup_file(file_path)
        except Exception as exc:
            logger.exception('Failed to process magic order %s', order_id)
            await self.orders_repo.mark_failed(order_id, error=str(exc))
            if str(exc) == 'BALANCE_TOO_LOW':
                await self._notify_failure(
                    bot,
                    telegram_id,
                    status_chat_id=status_chat_id,
                    status_message_id=status_message_id,
                    text=magic_order_balance_missing_text(price_uzs),
                )
            else:
                await self._notify_failure(
                    bot,
                    telegram_id,
                    status_chat_id=status_chat_id,
                    status_message_id=status_message_id,
                    text=magic_order_failed_text(),
                )
            if file_path:
                self._cleanup_file(file_path)

    async def _push_progress(
        self,
        *,
        bot: Bot,
        order_id,
        fallback_chat_id: int,
        payload: dict,
        percent: int,
        stage_key: str,
        status_chat_id: int | None,
        status_message_id: int | None,
        override_text: str | None = None,
    ) -> tuple[int, int]:
        text = override_text or magic_order_progress_text(payload=payload, percent=percent, stage_key=stage_key)
        if status_chat_id and status_message_id:
            try:
                await bot.edit_message_text(
                    chat_id=status_chat_id,
                    message_id=status_message_id,
                    text=text,
                )
                return int(status_chat_id), int(status_message_id)
            except TelegramBadRequest:
                pass

        sent = await bot.send_message(chat_id=fallback_chat_id, text=text)
        await self.orders_repo.set_status_message(order_id, chat_id=sent.chat.id, message_id=sent.message_id)
        return int(sent.chat.id), int(sent.message_id)

    async def _notify_failure(
        self,
        bot: Bot,
        telegram_id: int,
        *,
        status_chat_id: int | None,
        status_message_id: int | None,
        text: str,
    ) -> None:
        try:
            if status_chat_id and status_message_id:
                await bot.edit_message_text(
                    chat_id=status_chat_id,
                    message_id=status_message_id,
                    text=text,
                    reply_markup=magic_account_keyboard(),
                )
            else:
                await bot.send_message(chat_id=telegram_id, text=text, reply_markup=magic_account_keyboard())
        except Exception:
            logger.exception('Failed to notify user %s about magic order failure.', telegram_id)

    @staticmethod
    def _cleanup_file(file_path: str) -> None:
        path = Path(file_path)
        try:
            if path.exists():
                path.unlink()
        except OSError:
            logger.warning('Could not remove temporary premium file: %s', file_path)
