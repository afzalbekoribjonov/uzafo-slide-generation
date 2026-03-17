from __future__ import annotations

from typing import Literal

from pydantic import BaseModel, ConfigDict, Field, field_validator, model_validator


SectionContentType = Literal['facts', 'process', 'table']


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


PRESENTATION_PLAN_RESPONSE_SCHEMA = {
    'type': 'object',
    'description': 'Topic-grounded factual material for building a PowerPoint deck. The model must provide facts about the topic itself, not describe how the deck is organized.',
    'properties': {
        'presentation_title': {
            'type': 'string',
            'description': 'Short, topic-centered title. It must name the subject directly and avoid generic words like overview, presentation, deck, or automatically generated.'
        },
        'title_subtitle': {
            'type': 'string',
            'description': 'One concise subtitle that states the scope of the topic in factual terms. It must not mention slides, deck structure, or the act of presenting.'
        },
        'agenda_items': {
            'type': 'array',
            'minItems': 4,
            'maxItems': 8,
            'description': 'Core topic sections or historical periods that will be covered. Each item must be directly about the subject matter.',
            'items': {'type': 'string'},
        },
        'sections': {
            'type': 'array',
            'minItems': 3,
            'maxItems': 12,
            'description': 'Topic sections containing factual material. These are not instructions about slides; they are the actual facts to teach.',
            'items': {
                'type': 'object',
                'properties': {
                    'content_type': {
                        'type': 'string',
                        'enum': ['facts', 'process', 'table'],
                        'description': 'Use facts for general factual sections, process for chronology or step-by-step developments, and table for comparisons or dated summaries.'
                    },
                    'title': {
                        'type': 'string',
                        'description': 'Section title directly related to the topic, such as a period, dynasty, figure, reform, or concept.'
                    },
                    'focus': {
                        'type': 'string',
                        'description': 'One sentence explaining the factual lens of the section. This must be about the topic itself, not about presenting it.'
                    },
                    'facts': {
                        'type': 'array',
                        'minItems': 2,
                        'maxItems': 6,
                        'description': 'Concrete facts, dated developments, causes, results, characteristics, or examples tied to the section title.',
                        'items': {'type': 'string'},
                    },
                    'table': {
                        'type': 'object',
                        'description': 'Optional factual table used only when comparison or structured chronology is useful.',
                        'properties': {
                            'columns': {
                                'type': 'array',
                                'minItems': 2,
                                'maxItems': 5,
                                'items': {'type': 'string'},
                            },
                            'rows': {
                                'type': 'array',
                                'minItems': 2,
                                'maxItems': 6,
                                'items': {
                                    'type': 'array',
                                    'minItems': 2,
                                    'maxItems': 5,
                                    'items': {'type': 'string'},
                                },
                            },
                        },
                        'required': ['columns', 'rows'],
                    },
                },
                'required': ['content_type', 'title', 'focus', 'facts'],
            },
        },
        'summary_points': {
            'type': 'array',
            'minItems': 3,
            'maxItems': 5,
            'description': 'Final factual conclusions or takeaways about the topic. These must not thank the audience or mention AI or a presentation.',
            'items': {'type': 'string'},
        },
    },
    'required': ['presentation_title', 'title_subtitle', 'agenda_items', 'sections', 'summary_points'],
}
