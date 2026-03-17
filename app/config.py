from functools import lru_cache

from pydantic import AliasChoices, Field, field_validator
from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    bot_token: str = Field(alias='BOT_TOKEN')
    mongodb_uri: str = Field(alias='MONGODB_URI')
    mongodb_db: str = Field(default='slide_bot', alias='MONGODB_DB')
    admins: list[int] = Field(default_factory=list, validation_alias=AliasChoices('ADMINS', 'ADMIN_IDS'))
    support_contact: str = Field(default='@admin_username', alias='SUPPORT_CONTACT')
    generation_worker_poll_seconds: int = Field(default=3, alias='GENERATION_WORKER_POLL_SECONDS')
    generation_start_cooldown_seconds: int = Field(default=65, alias='GENERATION_START_COOLDOWN_SECONDS')
    gemini_api_key: str | None = Field(default=None, alias='GEMINI_API_KEY')
    gemini_model: str = Field(default='gemini-2.5-flash', alias='GEMINI_MODEL')
    gemini_max_retries: int = Field(default=3, alias='GEMINI_MAX_RETRIES')
    gemini_initial_backoff_seconds: int = Field(default=10, alias='GEMINI_INITIAL_BACKOFF_SECONDS')

    app_mode: str = Field(default='polling', alias='APP_MODE')
    web_server_host: str = Field(default='0.0.0.0', alias='WEB_SERVER_HOST')
    web_server_port: int = Field(default=10000, validation_alias=AliasChoices('PORT', 'WEB_SERVER_PORT'))
    webhook_base_url: str | None = Field(default=None, validation_alias=AliasChoices('WEBHOOK_BASE_URL', 'RENDER_EXTERNAL_URL'))
    webhook_path: str = Field(default='/telegram/webhook', alias='WEBHOOK_PATH')
    webhook_secret: str | None = Field(default=None, alias='WEBHOOK_SECRET')
    webhook_max_connections: int = Field(default=40, alias='WEBHOOK_MAX_CONNECTIONS')
    webhook_drop_pending_updates: bool = Field(default=False, alias='WEBHOOK_DROP_PENDING_UPDATES')

    model_config = SettingsConfigDict(
        env_file='.env',
        env_file_encoding='utf-8',
        case_sensitive=False,
        extra='ignore',
    )

    @field_validator('admins', mode='before')
    @classmethod
    def parse_admin_ids(cls, value):
        if value is None or value == '':
            return []

        if isinstance(value, int):
            return [value]

        if isinstance(value, list):
            return [int(x) for x in value]

        if isinstance(value, str):
            value = value.strip()
            if not value:
                return []
            return [int(x.strip()) for x in value.split(',') if x.strip()]

        return []

    @field_validator('app_mode', mode='before')
    @classmethod
    def normalize_app_mode(cls, value):
        if value is None:
            return 'polling'
        return str(value).strip().lower() or 'polling'

    @field_validator('webhook_base_url', mode='before')
    @classmethod
    def normalize_webhook_base_url(cls, value):
        if value is None:
            return None
        value = str(value).strip()
        return value.rstrip('/') or None

    @field_validator('webhook_path', mode='before')
    @classmethod
    def normalize_webhook_path(cls, value):
        value = str(value or '/telegram/webhook').strip()
        if not value.startswith('/'):
            value = f'/{value}'
        return value

    @property
    def webhook_url(self) -> str:
        if not self.webhook_base_url:
            raise ValueError('WEBHOOK_BASE_URL (yoki RENDER_EXTERNAL_URL) webhook rejimi uchun majburiy.')
        return f'{self.webhook_base_url}{self.webhook_path}'


@lru_cache
def get_settings() -> Settings:
    return Settings()
