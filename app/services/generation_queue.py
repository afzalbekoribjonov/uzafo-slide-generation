from __future__ import annotations

import asyncio
import logging
from pathlib import Path

from aiogram import Bot
from aiogram.exceptions import TelegramBadRequest
from aiogram.types import FSInputFile

from app.repositories.generations import GenerationsRepository
from app.repositories.users import UsersRepository
from app.services.pptx_generation import PptxGenerationService
from app.texts.user import (
    create_generation_failed_text,
    create_generation_progress_text,
    create_generation_success_caption,
)

logger = logging.getLogger(__name__)


class GenerationQueueService:
    def __init__(
        self,
        *,
        generations_repo: GenerationsRepository,
        users_repo: UsersRepository,
        pptx_generation_service: PptxGenerationService,
        poll_interval_seconds: int = 3,
        start_cooldown_seconds: int = 65,
    ) -> None:
        self.generations_repo = generations_repo
        self.users_repo = users_repo
        self.pptx_generation_service = pptx_generation_service
        self.poll_interval_seconds = max(1, int(poll_interval_seconds))
        self.start_cooldown_seconds = max(0, int(start_cooldown_seconds))
        self._last_started_monotonic: float | None = None

    async def create_job(
        self,
        *,
        telegram_id: int,
        full_name: str,
        username: str | None,
        payload: dict,
        consumed_from: str,
        status_chat_id: int | None = None,
        status_message_id: int | None = None,
    ) -> tuple[dict | None, int, dict | None]:
        active_job = await self.generations_repo.get_active_job_for_user(telegram_id)
        if active_job:
            ahead_count = await self.generations_repo.count_ahead_in_queue(active_job['_id'])
            return None, ahead_count, active_job

        job = await self.generations_repo.create_job(
            telegram_id=telegram_id,
            full_name=full_name,
            username=username,
            payload=payload,
            consumed_from=consumed_from,
            status_chat_id=status_chat_id,
            status_message_id=status_message_id,
        )
        ahead_count = await self.generations_repo.count_ahead_in_queue(job['_id'])
        return job, ahead_count, None

    async def run_worker(self, bot: Bot) -> None:
        requeued = await self.generations_repo.requeue_processing_jobs()
        if requeued:
            logger.warning('Requeued %s processing jobs after restart.', requeued)

        while True:
            try:
                await self._respect_cooldown()
                job = await self.generations_repo.claim_next_queued_job()
                if not job:
                    await asyncio.sleep(self.poll_interval_seconds)
                    continue

                self._last_started_monotonic = asyncio.get_running_loop().time()
                await self._process_job(bot, job)
            except asyncio.CancelledError:
                raise
            except Exception:
                logger.exception('Generation worker loop failed unexpectedly.')
                await asyncio.sleep(self.poll_interval_seconds)

    async def _respect_cooldown(self) -> None:
        if self._last_started_monotonic is None or self.start_cooldown_seconds <= 0:
            return

        elapsed = asyncio.get_running_loop().time() - self._last_started_monotonic
        remaining = self.start_cooldown_seconds - elapsed
        if remaining > 0:
            await asyncio.sleep(remaining)

    async def _process_job(self, bot: Bot, job: dict) -> None:
        generation_id = job['_id']
        telegram_id = int(job['telegram_id'])
        payload = dict(job.get('payload') or {})
        consumed_from = str(job.get('consumed_from') or 'free')
        file_path: str | None = None
        status_chat_id = job.get('status_chat_id')
        status_message_id = job.get('status_message_id')

        try:
            status_chat_id, status_message_id = await self._push_progress(
                bot=bot,
                generation_id=generation_id,
                fallback_chat_id=telegram_id,
                payload=payload,
                percent=12,
                stage_key='queued',
                status_chat_id=status_chat_id,
                status_message_id=status_message_id,
            )
            status_chat_id, status_message_id = await self._push_progress(
                bot=bot,
                generation_id=generation_id,
                fallback_chat_id=telegram_id,
                payload=payload,
                percent=34,
                stage_key='research',
                status_chat_id=status_chat_id,
                status_message_id=status_message_id,
            )
            plan = await asyncio.to_thread(self.pptx_generation_service.build_plan, payload)
            status_chat_id, status_message_id = await self._push_progress(
                bot=bot,
                generation_id=generation_id,
                fallback_chat_id=telegram_id,
                payload=payload,
                percent=61,
                stage_key='planning',
                status_chat_id=status_chat_id,
                status_message_id=status_message_id,
            )
            file_path = await asyncio.to_thread(self.pptx_generation_service.render, payload, plan)
            status_chat_id, status_message_id = await self._push_progress(
                bot=bot,
                generation_id=generation_id,
                fallback_chat_id=telegram_id,
                payload=payload,
                percent=82,
                stage_key='rendering',
                status_chat_id=status_chat_id,
                status_message_id=status_message_id,
            )
            status_chat_id, status_message_id = await self._push_progress(
                bot=bot,
                generation_id=generation_id,
                fallback_chat_id=telegram_id,
                payload=payload,
                percent=94,
                stage_key='uploading',
                status_chat_id=status_chat_id,
                status_message_id=status_message_id,
            )
            document = FSInputFile(file_path)
            await bot.send_document(
                chat_id=telegram_id,
                document=document,
                caption=create_generation_success_caption(payload),
            )
            await self.generations_repo.mark_done(generation_id, result_file_path=file_path)
            await self.users_repo.increment_successful_generation(telegram_id)
            await self._push_progress(
                bot=bot,
                generation_id=generation_id,
                fallback_chat_id=telegram_id,
                payload=payload,
                percent=100,
                stage_key='done',
                status_chat_id=status_chat_id,
                status_message_id=status_message_id,
            )
            self._cleanup_file(file_path)
        except Exception as exc:
            logger.exception('Failed to process generation job %s', generation_id)
            await self.generations_repo.mark_failed(generation_id, error=str(exc))
            await self.users_repo.restore_generation_credit(telegram_id, consumed_from)
            await self._notify_failure(
                bot,
                telegram_id,
                status_chat_id=status_chat_id,
                status_message_id=status_message_id,
            )
            if file_path:
                self._cleanup_file(file_path)

    async def _push_progress(
        self,
        *,
        bot: Bot,
        generation_id,
        fallback_chat_id: int,
        payload: dict,
        percent: int,
        stage_key: str,
        status_chat_id: int | None,
        status_message_id: int | None,
    ) -> tuple[int, int]:
        text = create_generation_progress_text(payload, percent, stage_key)
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
        await self.generations_repo.set_status_message(
            generation_id,
            chat_id=sent.chat.id,
            message_id=sent.message_id,
        )
        return int(sent.chat.id), int(sent.message_id)

    async def _notify_failure(
        self,
        bot: Bot,
        telegram_id: int,
        *,
        status_chat_id: int | None,
        status_message_id: int | None,
    ) -> None:
        text = create_generation_failed_text()
        try:
            if status_chat_id and status_message_id:
                await bot.edit_message_text(
                    chat_id=status_chat_id,
                    message_id=status_message_id,
                    text=text,
                )
            else:
                await bot.send_message(chat_id=telegram_id, text=text)
        except Exception:
            logger.exception('Failed to notify user %s about failed generation.', telegram_id)

    @staticmethod
    def _cleanup_file(file_path: str) -> None:
        path = Path(file_path)
        try:
            if path.exists():
                path.unlink()
        except OSError:
            logger.warning('Could not remove temporary file: %s', file_path)

    async def describe_existing_job(self, telegram_id: int) -> tuple[dict | None, int]:
        job = await self.generations_repo.get_active_job_for_user(telegram_id)
        if not job:
            return None, 0
        ahead_count = await self.generations_repo.count_ahead_in_queue(job['_id'])
        return job, ahead_count
