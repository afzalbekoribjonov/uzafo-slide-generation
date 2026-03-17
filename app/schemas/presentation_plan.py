from __future__ import annotations

from typing import Literal

from pydantic import BaseModel, ConfigDict, Field, field_validator, model_validator


SectionContentType = Literal["facts", "process", "table"]


class FactTable(BaseModel):
    model_config = ConfigDict(extra='ignore', str_strip_whitespace=True)

    columns: list[str] = Field(min_length=2, max_length=5)
    rows: list[list[str]] = Field(min_length=2, max_length=6)

    @field_validator('columns')
    @classmethod
    def clean_columns(cls, value: list[str]) -> list[str]:
        cleaned = [item.strip() for item in value if item and item.strip()]
        if len(cleaned) < 2:
            raise ValueError('Jadval uchun kamida 2 ta ustun kerak.')
        return cleaned[:5]

    @field_validator('rows')
    @classmethod
    def clean_rows(cls, value: list[list[str]]) -> list[list[str]]:
        cleaned_rows: list[list[str]] = []
        for row in value:
            cleaned = [str(item).strip() for item in row if str(item).strip()]
            if cleaned:
                cleaned_rows.append(cleaned)
        if len(cleaned_rows) < 2:
            raise ValueError('Jadval uchun kamida 2 ta qator kerak.')
        return cleaned_rows[:6]

    @model_validator(mode='after')
    def align_rows(self) -> 'FactTable':
        column_count = len(self.columns)
        normalized: list[list[str]] = []
        for row in self.rows:
            padded = row[:column_count]
            if len(padded) < column_count:
                padded.extend(['—'] * (column_count - len(padded)))
            normalized.append(padded)
        self.rows = normalized
        return self


class TopicSection(BaseModel):
    model_config = ConfigDict(extra='ignore', str_strip_whitespace=True)

    content_type: SectionContentType
    title: str = Field(min_length=3, max_length=90)
    focus: str = Field(min_length=10, max_length=180)
    facts: list[str] = Field(default_factory=list, max_length=6)
    table: FactTable | None = None

    @field_validator('facts')
    @classmethod
    def clean_facts(cls, value: list[str]) -> list[str]:
        cleaned = [item.strip() for item in value if item and item.strip()]
        return cleaned[:6]

    @model_validator(mode='after')
    def validate_payload(self) -> 'TopicSection':
        if self.content_type == 'table':
            if self.table is None:
                raise ValueError('table turidagi section uchun table ma’lumotlari kerak.')
            if len(self.facts) < 2:
                raise ValueError('table turidagi section uchun kamida 2 ta izoh/fakt kerak.')
        elif self.content_type == 'process':
            if len(self.facts) < 4:
                raise ValueError('process turidagi section uchun kamida 4 ta bosqich kerak.')
            self.facts = self.facts[:5]
        else:
            if len(self.facts) < 4:
                raise ValueError('facts turidagi section uchun kamida 4 ta fakt kerak.')
        return self


class PresentationPlan(BaseModel):
    model_config = ConfigDict(extra='ignore', str_strip_whitespace=True)

    presentation_title: str = Field(min_length=3, max_length=120)
    title_subtitle: str = Field(min_length=10, max_length=180)
    agenda_items: list[str] = Field(min_length=4, max_length=8)
    sections: list[TopicSection] = Field(min_length=3, max_length=12)
    summary_points: list[str] = Field(min_length=3, max_length=5)

    @field_validator('agenda_items', 'summary_points')
    @classmethod
    def clean_string_lists(cls, value: list[str]) -> list[str]:
        cleaned = [item.strip() for item in value if item and item.strip()]
        if not cleaned:
            raise ValueError('Bo‘sh ro‘yxat yuborildi.')
        return cleaned

    @model_validator(mode='after')
    def ensure_lengths(self) -> 'PresentationPlan':
        self.agenda_items = self.agenda_items[:8]
        self.sections = self.sections[:12]
        self.summary_points = self.summary_points[:5]
        return self


# Gemini response_schema serving can reject deeply nested or overly constrained schemas.
# Keep this schema intentionally compact. Table content is inferred later from section facts
# when content_type == "table", so the serving schema does not include nested table rows.
PRESENTATION_PLAN_RESPONSE_SCHEMA = {
    'type': 'object',
    'properties': {
        'presentation_title': {'type': 'string'},
        'title_subtitle': {'type': 'string'},
        'agenda_items': {'type': 'array', 'items': {'type': 'string'}},
        'sections': {
            'type': 'array',
            'items': {
                'type': 'object',
                'properties': {
                    'content_type': {'type': 'string', 'enum': ['facts', 'process', 'table']},
                    'title': {'type': 'string'},
                    'focus': {'type': 'string'},
                    'facts': {'type': 'array', 'items': {'type': 'string'}},
                },
                'required': ['content_type', 'title', 'focus', 'facts'],
            },
        },
        'summary_points': {'type': 'array', 'items': {'type': 'string'}},
    },
    'required': ['presentation_title', 'title_subtitle', 'agenda_items', 'sections', 'summary_points'],
}
