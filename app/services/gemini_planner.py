from __future__ import annotations

import json
import logging
import re
import time
from collections import Counter
from typing import Any

from pydantic import BaseModel, ConfigDict, Field, field_validator, model_validator

from app.schemas.presentation_plan import FactTable, PresentationPlan, TopicSection

logger = logging.getLogger(__name__)

try:
    from google import genai
    from google.genai import types
except Exception:  # pragma: no cover
    genai = None
    types = None


class GeminiPlannerError(RuntimeError):
    pass


class ResearchSection(BaseModel):
    model_config = ConfigDict(extra='ignore', str_strip_whitespace=True)

    title: str = Field(min_length=3, max_length=90)
    focus: str = Field(min_length=12, max_length=180)
    facts: list[str] = Field(min_length=4, max_length=7)

    @field_validator('facts')
    @classmethod
    def clean_facts(cls, value: list[str]) -> list[str]:
        cleaned = [item.strip() for item in value if item and item.strip()]
        if len(cleaned) < 4:
            raise ValueError('Har bir section uchun kamida 4 ta fakt kerak.')
        return cleaned[:7]


class ResearchDossier(BaseModel):
    model_config = ConfigDict(extra='ignore', str_strip_whitespace=True)

    topic_title: str = Field(min_length=3, max_length=120)
    scope_summary: str = Field(min_length=30, max_length=420)
    key_terms: list[str] = Field(min_length=4, max_length=10)
    section_notes: list[ResearchSection] = Field(min_length=3, max_length=12)
    final_takeaways: list[str] = Field(min_length=3, max_length=5)

    @field_validator('key_terms', 'final_takeaways')
    @classmethod
    def clean_lists(cls, value: list[str]) -> list[str]:
        cleaned = [item.strip() for item in value if item and item.strip()]
        if not cleaned:
            raise ValueError('Bo‘sh ro‘yxat yuborildi.')
        return cleaned

    @model_validator(mode='after')
    def trim_lists(self) -> 'ResearchDossier':
        self.key_terms = self.key_terms[:10]
        self.section_notes = self.section_notes[:12]
        self.final_takeaways = self.final_takeaways[:5]
        return self


RESEARCH_DOSSIER_RESPONSE_SCHEMA: dict[str, Any] = {
    'type': 'object',
    'properties': {
        'topic_title': {'type': 'string'},
        'scope_summary': {'type': 'string'},
        'key_terms': {'type': 'array', 'items': {'type': 'string'}},
        'section_notes': {
            'type': 'array',
            'items': {
                'type': 'object',
                'properties': {
                    'title': {'type': 'string'},
                    'focus': {'type': 'string'},
                    'facts': {'type': 'array', 'items': {'type': 'string'}},
                },
                'required': ['title', 'focus', 'facts'],
            },
        },
        'final_takeaways': {'type': 'array', 'items': {'type': 'string'}},
    },
    'required': ['topic_title', 'scope_summary', 'key_terms', 'section_notes', 'final_takeaways'],
}


class GeminiPresentationPlanner:
    def __init__(
        self,
        *,
        api_key: str | None,
        model_name: str = 'gemini-2.5-flash',
        max_retries: int = 3,
        initial_backoff_seconds: int = 10,
        fallback_models: list[str] | tuple[str, ...] | None = None,
    ) -> None:
        self.api_key = (api_key or '').strip()
        self.model_name = model_name.strip() or 'gemini-2.5-flash'
        self.max_retries = max(1, int(max_retries))
        self.initial_backoff_seconds = max(1, int(initial_backoff_seconds or 10))
        default_fallbacks = ('gemini-2.5-flash', 'gemini-2.0-flash')
        raw_fallbacks = fallback_models if fallback_models is not None else default_fallbacks
        self.fallback_models = [item.strip() for item in raw_fallbacks if item and item.strip()]

    @property
    def enabled(self) -> bool:
        return bool(self.api_key and genai is not None and types is not None)

    def build_plan(self, *, topic: str, presenter_name: str, slide_count: int, language_code: str) -> PresentationPlan:
        if not self.api_key:
            raise GeminiPlannerError('GEMINI_API_KEY topilmadi.')
        if genai is None or types is None:
            raise GeminiPlannerError('google-genai kutubxonasi o‘rnatilmagan.')

        section_count = max(3, min(12, slide_count - 3))
        client = genai.Client(api_key=self.api_key)

        dossier = self._run_json_call(
            client=client,
            prompt=self._build_research_prompt(
                topic=topic,
                presenter_name=presenter_name,
                slide_count=slide_count,
                section_count=section_count,
                language_code=language_code,
            ),
            response_schema=RESEARCH_DOSSIER_RESPONSE_SCHEMA,
            model=ResearchDossier,
            temperature=0.25,
            stage_name='research',
            use_server_schema=True,
            language_code=language_code,
            topic=topic,
        )

        try:
            plan = self._run_json_call(
                client=client,
                prompt=self._build_plan_prompt(
                    topic=topic,
                    presenter_name=presenter_name,
                    slide_count=slide_count,
                    section_count=section_count,
                    language_code=language_code,
                    dossier=dossier,
                ),
                response_schema=None,
                model=PresentationPlan,
                temperature=0.2,
                stage_name='planning',
                use_server_schema=False,
                language_code=language_code,
                topic=topic,
            )
            self._validate_plan_counts(plan=plan, expected_section_count=section_count)
            self._validate_plan_relevance(plan=plan, topic=topic)
            self._validate_plan_quality(plan=plan)
            return plan
        except Exception as exc:
            logger.warning('Gemini planning output was not usable, falling back to deterministic mapping: %s', exc)
            return self._convert_dossier_to_plan(
                topic=topic,
                dossier=dossier,
                section_count=section_count,
                language_code=language_code,
            )

    def _run_json_call(
        self,
        *,
        client: Any,
        prompt: str,
        response_schema: dict[str, Any] | None,
        model: type[BaseModel],
        temperature: float,
        stage_name: str,
        use_server_schema: bool = True,
        language_code: str = 'uz',
        topic: str = '',
    ) -> BaseModel:
        last_error: Exception | None = None

        for current_model in self._candidate_models():
            for attempt in range(1, self.max_retries + 1):
                try:
                    config_kwargs: dict[str, Any] = {
                        'temperature': temperature,
                        'top_p': 0.9,
                        'response_mime_type': 'application/json',
                    }
                    if use_server_schema and response_schema:
                        config_kwargs['response_schema'] = response_schema

                    response = client.models.generate_content(
                        model=current_model,
                        contents=prompt,
                        config=types.GenerateContentConfig(**config_kwargs),
                    )
                    raw_text = (getattr(response, 'text', None) or '').strip()
                    if not raw_text:
                        raise GeminiPlannerError(f'Gemini {stage_name} bosqichida bo‘sh javob qaytardi.')

                    data = json.loads(self._extract_json(raw_text))
                    if not isinstance(data, dict):
                        raise GeminiPlannerError(f'Gemini {stage_name} bosqichida JSON object qaytarmadi.')

                    normalized_data = self._normalize_payload_for_model(
                        data=data,
                        model=model,
                        language_code=language_code,
                        topic=topic,
                    )
                    return model.model_validate(normalized_data)
                except Exception as exc:  # pragma: no cover
                    last_error = exc
                    is_transient = self._is_transient_error(exc)
                    if attempt >= self.max_retries or not is_transient:
                        break
                    backoff_seconds = self.initial_backoff_seconds * (2 ** (attempt - 1))
                    logger.warning(
                        'Gemini %s request failed with model %s on attempt %s/%s. Retrying in %ss. Error: %s',
                        stage_name,
                        current_model,
                        attempt,
                        self.max_retries,
                        backoff_seconds,
                        exc,
                    )
                    time.sleep(backoff_seconds)

            if last_error is not None and self._is_transient_error(last_error):
                logger.warning(
                    'Gemini %s stage failed with model %s. Trying next candidate if available. Error: %s',
                    stage_name,
                    current_model,
                    last_error,
                )
                continue
            break

        raise GeminiPlannerError(f'Gemini {stage_name} bosqichi muvaffaqiyatsiz yakunlandi: {last_error}')


    def _convert_dossier_to_plan(
        self,
        *,
        topic: str,
        dossier: ResearchDossier,
        section_count: int,
        language_code: str,
    ) -> PresentationPlan:
        sections: list[TopicSection] = []
        notes = dossier.section_notes[:section_count]
        tables_used = 0
        table_slots = {max(2, section_count // 2), max(2, section_count - 1)}

        for index, note in enumerate(notes, start=1):
            combined = f'{note.title.lower()} {note.focus.lower()}'
            content_type = 'facts'
            table: FactTable | None = None
            facts = self._normalize_string_list(note.facts, max_items=6, item_max_length=160)

            if self._looks_process(combined):
                content_type = 'process'
                facts = self._ensure_fact_count(facts[:5], min_items=4, fallback_seed=note.focus, language_code=language_code)
            else:
                table_candidate = self._build_table_from_facts(
                    title=note.title,
                    focus=note.focus,
                    facts=note.facts,
                    language_code=language_code,
                )
                wants_table = self._looks_comparative(combined) or index in table_slots
                if table_candidate is not None and wants_table and tables_used < 2:
                    content_type = 'table'
                    table = table_candidate
                    tables_used += 1
                    facts = self._ensure_fact_count(facts[:2], min_items=2, fallback_seed=note.focus, language_code=language_code)
                else:
                    facts = self._ensure_fact_count(facts[:6], min_items=4, fallback_seed=note.focus, language_code=language_code)

            sections.append(
                TopicSection(
                    content_type=content_type,
                    title=self._fit_text(note.title, 90),
                    focus=self._fit_text(note.focus, 180),
                    facts=facts,
                    table=table,
                )
            )

        if tables_used == 0:
            for idx, section in enumerate(sections):
                if section.content_type != 'facts':
                    continue
                fallback_table = self._build_table_from_facts(
                    title=section.title,
                    focus=section.focus,
                    facts=section.facts,
                    language_code=language_code,
                )
                if fallback_table is not None and self._looks_comparative(f'{section.title} {section.focus}'):
                    sections[idx] = TopicSection(
                        content_type='table',
                        title=section.title,
                        focus=section.focus,
                        facts=section.facts[:2],
                        table=fallback_table,
                    )
                    tables_used += 1
                    break

        agenda_items = self._normalize_string_list([note.title for note in notes], max_items=8, item_max_length=70)
        if len(agenda_items) < 4:
            agenda_items.extend(
                self._fallback_agenda_items(
                    topic=topic,
                    existing=agenda_items,
                    needed=4 - len(agenda_items),
                    language_code=language_code,
                )
            )

        summary_points = self._normalize_string_list(dossier.final_takeaways, max_items=5, item_max_length=170)
        if len(summary_points) < 3:
            summary_points.extend(
                self._fallback_summary_points(
                    topic=topic,
                    needed=3 - len(summary_points),
                    language_code=language_code,
                )
            )

        return PresentationPlan(
            presentation_title=self._fit_text(topic or dossier.topic_title, 120),
            title_subtitle=self._fit_text(dossier.scope_summary, 180),
            agenda_items=agenda_items[:8],
            sections=sections,
            summary_points=summary_points[:5],
        )

    @staticmethod
    def _extract_json(raw_text: str) -> str:
        text = raw_text.strip()
        if text.startswith('```'):
            text = re.sub(r'^```(?:json)?\s*', '', text)
            text = re.sub(r'\s*```$', '', text)
        return text.strip()

    @staticmethod
    def _is_transient_error(exc: Exception) -> bool:
        message = str(exc).upper()
        transient_markers = ('429', '503', 'RESOURCE_EXHAUSTED', 'UNAVAILABLE', 'TIMEOUT', 'DEADLINE')
        return any(marker in message for marker in transient_markers)

    @staticmethod
    def _validate_plan_counts(*, plan: PresentationPlan, expected_section_count: int) -> None:
        actual = len(plan.sections)
        if actual != expected_section_count:
            raise GeminiPlannerError(
                f'Gemini noto‘g‘ri section soni qaytardi. Kutilgan: {expected_section_count}, kelgan: {actual}.'
            )

    @staticmethod
    def _validate_plan_relevance(*, plan: PresentationPlan, topic: str) -> None:
        blocked_markers = (
            'taqdimot qanday', 'presentation structure', 'deck covers', 'automatically generated',
            'generated automatically', 'e’tiboringiz uchun rahmat', 'thank you for your attention',
            'slides', 'slide deck', 'this presentation', 'taqdimot tuzilmasi', 'what this deck covers',
        )
        all_text_parts = [plan.presentation_title, plan.title_subtitle, *plan.agenda_items, *plan.summary_points]
        for section in plan.sections:
            all_text_parts.extend([section.title, section.focus, *section.facts])
            if section.table:
                all_text_parts.extend(section.table.columns)
                for row in section.table.rows:
                    all_text_parts.extend(row)
        blob = ' '.join(all_text_parts).lower()
        if any(marker in blob for marker in blocked_markers):
            raise GeminiPlannerError('Gemini meta-matn qaytardi, plan rad etildi.')

        keywords = GeminiPresentationPlanner._topic_keywords(topic)
        if keywords and not any(keyword in blob for keyword in keywords):
            raise GeminiPlannerError('Gemini javobi mavzu bilan yetarli darajada bog‘lanmadi.')

    @staticmethod
    def _validate_plan_quality(*, plan: PresentationPlan) -> None:
        generic_starts = (
            'bu bo‘lim', 'ushbu bo‘lim', 'the section', 'this section', 'mavzu muhim', 'topic is important',
            'mazmun umumiy', 'taqdimot', 'mavzuni yoritadi',
        )
        too_generic = 0
        for section in plan.sections:
            if section.focus.lower().startswith(generic_starts):
                too_generic += 1
            for fact in section.facts:
                lowered = fact.lower()
                if lowered.startswith(generic_starts) or len(lowered.split()) < 5:
                    too_generic += 1
        if too_generic >= 4:
            raise GeminiPlannerError('Gemini juda ko‘p umumiy/generic matn qaytardi.')

    @staticmethod
    def _topic_keywords(topic: str) -> list[str]:
        stopwords = {
            'va', 'the', 'of', 'in', 'on', 'for', 'to', 'a', 'an', 'tarixi', 'history', 'madaniyati', 'culture',
            'haqida', 'about', 'asoslari', 'introduction', 'kirish',
        }
        parts = re.findall(r"[A-Za-zÀ-ÿА-Яа-я0-9ʻ’'\-]+", topic.lower())
        keywords = [part for part in parts if len(part) > 2 and part not in stopwords]
        return keywords[:6]

    @staticmethod
    def _language_name(language_code: str) -> str:
        return {
            'uz': 'Uzbek',
            'ru': 'Russian',
            'en': 'English',
        }.get(language_code, 'Uzbek')


    @staticmethod
    def _language_pack(language_code: str) -> dict[str, Any]:
        if language_code == 'ru':
            return {
                'unknown_topic': 'Тема',
                'section_label': 'Раздел',
                'scope_template': '{topic}: кратко раскрываются основные понятия, структура и важные особенности темы.',
                'focus_template': '{title} раскрывает ключевые аспекты темы {topic}.',
                'focus_section_template': '{title} помогает последовательно объяснить ключевое содержание темы.',
                'fact_pool': [
                    '{base}',
                    '{base} объясняется через исторический, научный или практический контекст.',
                    '{base} рассматривается через причины, признаки и последствия.',
                    '{base} помогает выделить наиболее важные свойства темы.',
                    '{base} показывает значимость темы на практике или в теории.',
                ],
                'key_terms_seed': ['ключевое понятие', 'развитие', 'особенность', 'значение'],
                'agenda_seed': ['Введение', 'Основные понятия', 'Ключевые особенности', 'Практическое значение', 'Вывод'],
                'agenda_suffix': '{topic}: основной раздел',
                'summary_seed': [
                    '{topic} раскрывается через несколько взаимосвязанных аспектов.',
                    'Содержание темы важно рассматривать через причины, развитие и последствия.',
                    'Знание этой темы помогает глубже понять её значение и применение.',
                    'Анализ темы показывает связи между её основными понятиями и примерами.',
                ],
                'generic_bad_headers': {'основной аспект', 'описание', 'важный признак', 'заметка'},
                'comparison_headers': ('Аспект', 'Ключевая деталь'),
            }
        if language_code == 'en':
            return {
                'unknown_topic': 'Topic',
                'section_label': 'Section',
                'scope_template': '{topic}: the core concepts, structure, and most important characteristics are summarized concisely.',
                'focus_template': '{title} explains key aspects of the topic {topic}.',
                'focus_section_template': '{title} helps explain the core content of the topic in a clear sequence.',
                'fact_pool': [
                    '{base}',
                    '{base} is explained through historical, scientific, or practical context.',
                    '{base} is examined through causes, features, and consequences.',
                    '{base} highlights the most important characteristics of the topic.',
                    '{base} shows why the topic matters in practice or theory.',
                ],
                'key_terms_seed': ['core concept', 'development', 'feature', 'importance'],
                'agenda_seed': ['Introduction', 'Core concepts', 'Key features', 'Practical relevance', 'Conclusion'],
                'agenda_suffix': '{topic}: main section',
                'summary_seed': [
                    '{topic} is understood most clearly through its interconnected dimensions.',
                    'The topic becomes clearer when its causes, development, and outcomes are connected.',
                    'Knowledge of the topic helps explain its importance and application.',
                    'A focused analysis reveals the links between the topic’s main concepts and examples.',
                ],
                'generic_bad_headers': {'main aspect', 'description', 'important feature', 'note'},
                'comparison_headers': ('Aspect', 'Key detail'),
            }
        return {
            'unknown_topic': 'Mavzu',
            'section_label': 'Bo‘lim',
            'scope_template': '{topic}: mavzuning asosiy tushunchalari, tuzilishi va muhim xususiyatlari qisqacha yoritiladi.',
            'focus_template': '{title} {topic} mavzusining muhim jihatlarini tushuntiradi.',
            'focus_section_template': '{title} mavzuning asosiy mazmunini izchil yoritishga yordam beradi.',
            'fact_pool': [
                '{base}',
                '{base} tarixiy, ilmiy yoki amaliy kontekst orqali tushuntiriladi.',
                '{base} sabablar, belgilar va natijalar orqali ko‘rib chiqiladi.',
                '{base} mavzuning eng muhim xususiyatlarini ajratib ko‘rsatadi.',
                '{base} mavzuning amaliy yoki nazariy ahamiyatini ko‘rsatadi.',
            ],
            'key_terms_seed': ['asosiy tushuncha', 'rivojlanish', 'xususiyat', 'ahamiyat'],
            'agenda_seed': ['Kirish', 'Asosiy tushunchalar', 'Muhim xususiyatlar', 'Amaliy ahamiyati', 'Xulosa'],
            'agenda_suffix': '{topic}: asosiy bo‘lim',
            'summary_seed': [
                '{topic} bir nechta o‘zaro bog‘liq jihatlar orqali tushuniladi.',
                'Mavzuni sabab, rivojlanish va natijalar bilan birga ko‘rish mazmunni ravshanlashtiradi.',
                'Mavzu haqidagi bilim uning ahamiyati va qo‘llanilishini chuqurroq anglashga yordam beradi.',
                'Tahlil mavzuning asosiy tushunchalari va misollari o‘rtasidagi bog‘liqlikni ochib beradi.',
            ],
            'generic_bad_headers': {'asosiy jihat', 'tavsif', 'muhim belgi', 'izoh'},
            'comparison_headers': ('Jihat', 'Asosiy tafsilot'),
        }

    def _build_research_prompt(
        self,
        *,
        topic: str,
        presenter_name: str,
        slide_count: int,
        section_count: int,
        language_code: str,
    ) -> str:
        language_name = self._language_name(language_code)
        domain_rules = self._domain_rules(topic)
        return f"""You are a meticulous {self._domain_role(topic)} writing ONLY in {language_name}. Your task is to collect factual notes about the topic itself. Do not mention slides, presentation structure, audience, or the act of presenting. Return only valid JSON matching the requested keys.

Topic: {topic}
Presenter name for attribution only: {presenter_name}
Total requested slides: {slide_count}
Required factual sections: {section_count}

Research rules:
- Every output string must be in {language_name}.
- Every sentence must teach the topic itself.
- scope_summary must be a single concise sentence and no longer than 170 characters.
- Use concrete names, places, dates, developments, causes, consequences, and distinguishing features when relevant.
- Avoid filler, introductions about the deck, motivational language, and generic presentation phrases.
- Build exactly the requested number of section_notes.
- Each section must contain 4 to 7 concrete facts.
- Each fact must be one complete sentence, concise but meaningful, ideally suited for a single slide bullet.
- final_takeaways must be factual conclusions, not thanks or closing ceremony language.
- If uncertain, stay with well-established facts and avoid invented detail.

Domain-specific guidance:
{domain_rules}
"""


    def _build_plan_prompt(
        self,
        *,
        topic: str,
        presenter_name: str,
        slide_count: int,
        section_count: int,
        language_code: str,
        dossier: ResearchDossier,
    ) -> str:
        language_name = self._language_name(language_code)
        dossier_json = json.dumps(dossier.model_dump(mode='json'), ensure_ascii=False, indent=2)
        return f"""You are a careful content editor working ONLY in {language_name}. Transform the research dossier into a factual teaching plan. You are NOT a presentation writer. Do not talk about slides, the audience, or what the deck will do. Return only valid JSON as a plain object with these keys: presentation_title, title_subtitle, agenda_items, sections, summary_points.

Topic: {topic}
Presenter name for cover only: {presenter_name}
Total requested slides: {slide_count}
Required factual body sections: {section_count}

Strict rules:
- Every output string must be in {language_name}.
- Use only topic-grounded content.
- title_subtitle must be one sentence and no longer than 170 characters.
- agenda_items must be 4 to 8 short topic sections or periods.
- sections must equal the required body section count exactly.
- Every section object must contain: content_type, title, focus, facts.
- Vary the body layouts naturally across the deck when the content supports it. Prefer a mix of facts, process, and at least one table or comparison section when the topic genuinely contains structured or comparable information.
- Use content_type="process" only when chronology or sequence is clear.
- Use content_type="table" only when there is real comparative or structured data.
- Prefer 1 or 2 table sections total, never force more than 2.
- If a table is used, the table data must be carried inside section.table with keys columns and rows.
- Table columns must be concise, topic-specific labels in the chosen language.
- Table cell values must be short and compact, ideally 2 to 7 words.
- Table rows must be arrays of plain strings only. Never serialize a dictionary or object as a string inside a cell.
- Do not repeat the full section title in every table row.
- If structured table data is not clearly available, use content_type="facts" instead of forcing a table.
- Summary points must be one complete sentence each, concise and readable on a slide card.
- Never use phrases like "this presentation", "overview", "thank you", or "what this deck covers".

Research dossier:
{dossier_json}
"""

    @staticmethod
    def _domain_role(topic: str) -> str:
        lower = topic.lower()
        if any(marker in lower for marker in ('tarix', 'history', 'dynasty', 'culture', 'madaniyat', 'heritage')):
            return 'historian and cultural researcher'
        if any(marker in lower for marker in ('biology', 'kimyo', 'physics', 'science', 'fizika', 'biologiya')):
            return 'scientific explainer'
        if any(marker in lower for marker in ('business', 'marketing', 'iqtisod', 'economy', 'finance')):
            return 'business analyst'
        return 'subject-matter expert'

    @staticmethod
    def _domain_rules(topic: str) -> str:
        lower = topic.lower()
        if any(marker in lower for marker in ('tarix', 'history', 'culture', 'madaniyat', 'heritage')):
            return (
                '- Organize section_notes around real historical periods, states, rulers, cities, institutions, reforms, conflicts, cultural schools, and legacy.\n'
                '- Use proper names and approximate dates whenever they are widely known.\n'
                '- A table is useful for comparing periods, dynasties, cities, or contributions.\n'
            )
        if any(marker in lower for marker in ('business', 'marketing', 'economy', 'finance', 'iqtisod')):
            return (
                '- Focus on mechanisms, drivers, outcomes, metrics, market conditions, and examples.\n'
                '- A table is useful for comparisons, metrics, and scenario differences.\n'
            )
        return (
            '- Build section_notes around the main dimensions, examples, causes, outcomes, and practical implications of the topic.\n'
            '- Use a table only if it genuinely helps compare concepts or structured facts.\n'
        )


    @staticmethod
    def _looks_process(text: str) -> bool:
        lowered = text.lower()
        markers = (
            'davr', 'bosqich', 'jarayon', 'sequence', 'chronology', 'timeline', 'rivojlanish',
            'evolution', 'stage', 'этап', 'период', 'процесс',
        )
        return any(marker in lowered for marker in markers)

    @staticmethod
    def _looks_comparative(text: str) -> bool:
        lowered = text.lower()
        markers = (
            'taqqos', 'qiyos', 'turlari', 'turlar', 'klassifik', 'tasnif', 'comparison',
            'compare', 'types', 'type', 'categories', 'classification', 'matrix',
            'versus', 'vs', 'contrast', 'сравн', 'виды', 'типы', 'классиф', 'категории',
        )
        return any(marker in lowered for marker in markers)

    def _candidate_models(self) -> list[str]:
        seen: set[str] = set()
        candidates: list[str] = []
        for item in [self.model_name, *self.fallback_models]:
            normalized = (item or '').strip()
            if normalized and normalized not in seen:
                seen.add(normalized)
                candidates.append(normalized)
        return candidates or ['gemini-2.5-flash']

    def _normalize_payload_for_model(
        self,
        *,
        data: dict[str, Any],
        model: type[BaseModel],
        language_code: str,
        topic: str,
    ) -> dict[str, Any]:
        if model is ResearchDossier:
            return self._normalize_research_payload(data, language_code=language_code, topic=topic)
        if model is PresentationPlan:
            return self._normalize_plan_payload(data, language_code=language_code, topic=topic)
        return data

    def _normalize_research_payload(self, data: dict[str, Any], *, language_code: str, topic: str) -> dict[str, Any]:
        pack = self._language_pack(language_code)
        fallback_topic = self._fit_text(topic or pack['unknown_topic'], 120)
        topic_title = self._fit_text(data.get('topic_title') or data.get('presentation_title') or fallback_topic, 120)
        scope_summary = self._fit_text(
            data.get('scope_summary') or data.get('title_subtitle') or str(pack['scope_template']).format(topic=topic_title),
            420,
        )
        if len(scope_summary) < 30:
            scope_summary = self._fit_text(str(pack['scope_template']).format(topic=topic_title), 420)

        key_terms = self._normalize_string_list(data.get('key_terms'), max_items=10, item_max_length=60)
        if len(key_terms) < 4:
            key_terms.extend(
                self._fallback_key_terms(
                    topic_title=topic_title,
                    existing=key_terms,
                    needed=4 - len(key_terms),
                    language_code=language_code,
                )
            )

        raw_sections = data.get('section_notes') if isinstance(data.get('section_notes'), list) else []
        section_notes: list[dict[str, Any]] = []
        for index, raw_section in enumerate(raw_sections[:12], start=1):
            if not isinstance(raw_section, dict):
                continue
            title = self._fit_text(raw_section.get('title') or f"{pack['section_label']} {index}", 90)
            focus = self._fit_text(
                raw_section.get('focus') or str(pack['focus_template']).format(title=title, topic=topic_title),
                180,
            )
            if len(focus) < 12:
                focus = self._fit_text(str(pack['focus_section_template']).format(title=title, topic=topic_title), 180)

            facts = self._normalize_string_list(raw_section.get('facts'), max_items=7, item_max_length=170)
            facts = self._ensure_fact_count(facts, min_items=4, fallback_seed=f'{title}. {focus}', language_code=language_code)
            section_notes.append({'title': title, 'focus': focus, 'facts': facts[:7]})

        while len(section_notes) < 3:
            index = len(section_notes) + 1
            title = self._fit_text(f"{topic_title}: {pack['section_label']} {index}", 90)
            focus = self._fit_text(str(pack['focus_section_template']).format(title=title, topic=topic_title), 180)
            section_notes.append(
                {
                    'title': title,
                    'focus': focus,
                    'facts': self._ensure_fact_count([], min_items=4, fallback_seed=f'{title}. {focus}', language_code=language_code),
                }
            )

        final_takeaways = self._normalize_string_list(data.get('final_takeaways'), max_items=5, item_max_length=170)
        if len(final_takeaways) < 3:
            final_takeaways.extend(
                self._fallback_summary_points(
                    topic=topic_title,
                    needed=3 - len(final_takeaways),
                    language_code=language_code,
                )
            )

        return {
            'topic_title': topic_title,
            'scope_summary': scope_summary,
            'key_terms': key_terms[:10],
            'section_notes': section_notes[:12],
            'final_takeaways': final_takeaways[:5],
        }

    def _normalize_plan_payload(self, data: dict[str, Any], *, language_code: str, topic: str) -> dict[str, Any]:
        pack = self._language_pack(language_code)
        fallback_title = self._fit_text(topic or pack['unknown_topic'], 120)
        presentation_title = self._fit_text(data.get('presentation_title') or data.get('topic_title') or fallback_title, 120)
        title_subtitle = self._fit_text(
            data.get('title_subtitle') or data.get('scope_summary') or str(pack['scope_template']).format(topic=presentation_title),
            180,
        )
        if len(title_subtitle) < 10:
            title_subtitle = self._fit_text(str(pack['scope_template']).format(topic=presentation_title), 180)

        agenda_items = self._normalize_string_list(data.get('agenda_items'), max_items=8, item_max_length=70)

        raw_sections = data.get('sections') if isinstance(data.get('sections'), list) else []
        sections: list[dict[str, Any]] = []
        for index, raw_section in enumerate(raw_sections[:12], start=1):
            if not isinstance(raw_section, dict):
                continue

            content_type = str(raw_section.get('content_type') or 'facts').strip().lower()
            if content_type not in {'facts', 'process', 'table'}:
                content_type = 'facts'

            title = self._fit_text(raw_section.get('title') or f"{pack['section_label']} {index}", 90)
            focus = self._fit_text(
                raw_section.get('focus') or str(pack['focus_template']).format(title=title, topic=presentation_title),
                180,
            )
            facts = self._normalize_string_list(raw_section.get('facts'), max_items=6, item_max_length=160)

            section_payload: dict[str, Any] = {
                'content_type': content_type,
                'title': title,
                'focus': focus,
                'facts': facts,
            }

            if content_type == 'process':
                section_payload['facts'] = self._ensure_fact_count(facts[:5], min_items=4, fallback_seed=f'{title}. {focus}', language_code=language_code)
            elif content_type == 'table':
                normalized_table = self._normalize_table_candidate(
                    title=title,
                    focus=focus,
                    table_raw=raw_section.get('table') if isinstance(raw_section.get('table'), dict) else {},
                    source_facts=facts,
                    language_code=language_code,
                )
                if normalized_table is None:
                    section_payload['content_type'] = 'facts'
                    section_payload['facts'] = self._ensure_fact_count(facts, min_items=4, fallback_seed=f'{title}. {focus}', language_code=language_code)
                else:
                    section_payload['facts'] = self._ensure_fact_count(facts[:2], min_items=2, fallback_seed=f'{title}. {focus}', language_code=language_code)
                    section_payload['table'] = normalized_table.model_dump(mode='json')
            else:
                section_payload['facts'] = self._ensure_fact_count(facts, min_items=4, fallback_seed=f'{title}. {focus}', language_code=language_code)

            sections.append(section_payload)

        if len(agenda_items) < 4 and sections:
            agenda_items.extend(self._normalize_string_list([section['title'] for section in sections], max_items=8, item_max_length=70))
        if len(agenda_items) < 4:
            agenda_items.extend(
                self._fallback_agenda_items(
                    topic=presentation_title,
                    existing=agenda_items,
                    needed=4 - len(agenda_items),
                    language_code=language_code,
                )
            )

        summary_points = self._normalize_string_list(data.get('summary_points'), max_items=5, item_max_length=170)
        if len(summary_points) < 3:
            summary_points.extend(
                self._fallback_summary_points(
                    topic=presentation_title,
                    needed=3 - len(summary_points),
                    language_code=language_code,
                )
            )

        return {
            'presentation_title': presentation_title,
            'title_subtitle': title_subtitle,
            'agenda_items': agenda_items[:8],
            'sections': sections[:12],
            'summary_points': summary_points[:5],
        }

    def _fit_text(self, value: Any, max_length: int) -> str:
        text = self._clean_text_artifacts(value)
        if not text:
            return ''
        if len(text) <= max_length:
            return text
        shortened = text[: max_length + 1]
        last_space = shortened.rfind(' ')
        if last_space >= max_length * 0.6:
            shortened = shortened[:last_space]
        else:
            shortened = shortened[:max_length]
        return shortened.rstrip(' ,;:-.')

    @staticmethod
    def _split_top_level_parts(text: str) -> list[str]:
        parts: list[str] = []
        buffer: list[str] = []
        quote: str | None = None
        depth = 0
        for char in text:
            if char in {'"', "'"}:
                if quote is None:
                    quote = char
                elif quote == char:
                    quote = None
                buffer.append(char)
                continue
            if quote is None:
                if char in '{[(':
                    depth += 1
                elif char in '}])' and depth > 0:
                    depth -= 1
                elif char in {',', ';'} and depth == 0:
                    part = ''.join(buffer).strip()
                    if part:
                        parts.append(part)
                    buffer = []
                    continue
            buffer.append(char)
        tail = ''.join(buffer).strip()
        if tail:
            parts.append(tail)
        return parts

    @staticmethod
    def _strip_wrapping_quotes(text: str) -> str:
        cleaned = text.strip()
        quote_pairs = [('"', '"'), ("'", "'"), ('“', '”'), ('‘', '’')]
        changed = True
        while changed and len(cleaned) >= 2:
            changed = False
            for left, right in quote_pairs:
                if cleaned.startswith(left) and cleaned.endswith(right):
                    cleaned = cleaned[1:-1].strip()
                    changed = True
        return cleaned

    def _clean_text_artifacts(self, value: Any) -> str:
        text = '' if value is None else str(value)
        text = text.replace(' ', ' ')
        text = text.replace('\"', '"').replace("\'", "'")
        text = text.replace('“', '"').replace('”', '"').replace('‘', "'").replace('’', "'")
        text = re.sub(r'\s+', ' ', text).strip()
        if not text:
            return ''
        if (text.startswith('{') and text.endswith('}')) or (text.startswith('[') and text.endswith(']')):
            pairs = self._extract_key_value_pairs(text)
            if pairs:
                text = '; '.join(f'{key}: {val}' for key, val in pairs)
            else:
                text = text[1:-1].strip()
        text = self._strip_wrapping_quotes(text)
        text = text.replace('{', '').replace('}', '').replace('[', '').replace(']', '')
        text = re.sub(r'\s*([,:;])\s*', r'\1 ', text)
        text = re.sub(r'\s+', ' ', text).strip(" \"'\n\t-–—")
        return text

    def _normalize_string_list(self, value: Any, *, max_items: int, item_max_length: int) -> list[str]:
        if not isinstance(value, list):
            return []
        cleaned: list[str] = []
        for item in value:
            text = self._fit_text(item, item_max_length)
            if text:
                cleaned.append(text)
            if len(cleaned) >= max_items:
                break
        return cleaned

    def _ensure_fact_count(self, items: list[str], *, min_items: int, fallback_seed: str, language_code: str = 'uz') -> list[str]:
        cleaned = [self._fit_text(item, 170) for item in items if self._fit_text(item, 170)]
        fallback_pool = self._fallback_fact_pool(fallback_seed, language_code=language_code)
        pool_index = 0
        while len(cleaned) < min_items and pool_index < len(fallback_pool):
            candidate = fallback_pool[pool_index]
            pool_index += 1
            if candidate not in cleaned:
                cleaned.append(candidate)
        return cleaned

    def _fallback_fact_pool(self, seed: str, *, language_code: str = 'uz') -> list[str]:
        pack = self._language_pack(language_code)
        base = self._fit_text(seed, 160) or str(pack['unknown_topic'])
        return [self._fit_text(template.format(base=base), 170) for template in pack['fact_pool']]

    def _extract_key_value_pairs(self, value: Any) -> list[tuple[str, str]]:
        text = '' if value is None else str(value)
        text = text.replace(' ', ' ').replace('\"', '"').replace("\'", "'")
        text = text.strip()
        if not text or ':' not in text:
            return []
        if (text.startswith('{') and text.endswith('}')) or (text.startswith('[') and text.endswith(']')):
            text = text[1:-1].strip()
        parts = self._split_top_level_parts(text)
        pairs: list[tuple[str, str]] = []
        for part in parts:
            if ':' not in part:
                continue
            key, raw_val = part.split(':', 1)
            clean_key = self._clean_text_artifacts(key)
            clean_val = self._clean_text_artifacts(raw_val)
            if clean_key and clean_val:
                pairs.append((clean_key, clean_val))
        return pairs

    def _table_columns_look_generic(self, columns: list[str], language_code: str) -> bool:
        if not columns:
            return True
        bad = {item.lower() for item in self._language_pack(language_code)['generic_bad_headers']}
        normalized = [self._clean_text_artifacts(col).lower() for col in columns]
        return len(normalized) >= 3 and all(col in bad for col in normalized[:3])

    def _normalize_table_candidate(
        self,
        *,
        title: str,
        focus: str,
        table_raw: dict[str, Any],
        source_facts: list[str],
        language_code: str,
    ) -> FactTable | None:
        source_facts = self._normalize_string_list(source_facts, max_items=6, item_max_length=170)
        columns = self._normalize_string_list(table_raw.get('columns'), max_items=5, item_max_length=26)
        rows_raw = table_raw.get('rows') if isinstance(table_raw.get('rows'), list) else []
        rows: list[list[str]] = []

        if columns and rows_raw:
            for row in rows_raw[:6]:
                normalized_row = self._normalize_table_row(
                    row=row,
                    columns=columns,
                    row_index=len(rows) + 1,
                    title=title,
                    focus=focus,
                )
                if normalized_row and len(normalized_row) >= min(2, len(columns)):
                    padded = normalized_row[: len(columns)]
                    if len(padded) < len(columns):
                        padded.extend(['—'] * (len(columns) - len(padded)))
                    rows.append(padded)
            if len(rows) >= 2 and not self._table_columns_look_generic(columns, language_code):
                return FactTable(columns=columns, rows=rows[:6])

        inferred = self._build_table_from_facts(title=title, focus=focus, facts=source_facts, language_code=language_code)
        if inferred is not None:
            return inferred

        if columns and rows and len(rows) >= 2:
            return FactTable(columns=columns, rows=rows[:6])
        return None

    def _infer_table_columns(self, *, title: str, focus: str, facts: list[str], language_code: str) -> list[str]:
        key_counter: Counter[str] = Counter()
        first_seen: list[str] = []
        pretty_names: dict[str, str] = {}
        for fact in facts[:6]:
            pairs = self._extract_key_value_pairs(fact)
            seen_in_row: set[str] = set()
            for key, _ in pairs:
                normalized = self._clean_text_artifacts(key).lower()
                if not normalized:
                    continue
                if normalized not in pretty_names:
                    pretty_names[normalized] = self._fit_text(key, 26)
                    first_seen.append(normalized)
                if normalized not in seen_in_row:
                    key_counter[normalized] += 1
                    seen_in_row.add(normalized)
        ordered = sorted(
            [key for key, count in key_counter.items() if count >= 2],
            key=lambda key: (-key_counter[key], first_seen.index(key)),
        )
        columns = [pretty_names[key] for key in ordered[:4]]
        return columns if len(columns) >= 2 else []

    def _normalize_table_row(
        self,
        *,
        row: Any,
        columns: list[str],
        row_index: int,
        title: str,
        focus: str,
    ) -> list[str] | None:
        if isinstance(row, dict):
            lowered_map = {self._clean_text_artifacts(key).lower(): value for key, value in row.items()}
            normalized = [self._fit_text(lowered_map.get(column.lower(), '—'), 60) for column in columns]
            return normalized if any(cell != '—' for cell in normalized) else None

        if isinstance(row, list):
            cleaned_row = [self._fit_text(cell, 60) for cell in row[: len(columns)] if self._fit_text(cell, 60)]
            return cleaned_row if cleaned_row else None

        if isinstance(row, str):
            pairs = self._extract_key_value_pairs(row)
            if len(pairs) >= 2:
                return self._pairs_to_row(pairs, columns)
            return None

        return None

    def _pairs_to_row(self, pairs: list[tuple[str, str]], columns: list[str]) -> list[str]:
        values_by_key = {self._clean_text_artifacts(key).lower(): val for key, val in pairs}
        ordered_values = [self._fit_text(values_by_key.get(column.lower(), '—'), 60) for column in columns]
        if sum(1 for item in ordered_values if item != '—') >= min(2, len(columns)):
            return ordered_values[: len(columns)]
        raw_values = [self._fit_text(val, 60) for _, val in pairs[: len(columns)]]
        if len(raw_values) < len(columns):
            raw_values.extend(['—'] * (len(columns) - len(raw_values)))
        return raw_values[: len(columns)]


    def _build_table_from_facts(self, title: str, focus: str, facts: list[str], language_code: str) -> FactTable | None:
        normalized_facts = [self._fit_text(fact, 170) for fact in facts if self._fit_text(fact, 170)]
        columns = self._infer_table_columns(title=title, focus=focus, facts=normalized_facts, language_code=language_code)
        if len(columns) >= 2:
            rows: list[list[str]] = []
            for fact in normalized_facts[:6]:
                pairs = self._extract_key_value_pairs(fact)
                if len(pairs) >= 2:
                    row = self._pairs_to_row(pairs, columns)
                    if sum(1 for item in row if item != '—') >= min(2, len(columns)):
                        rows.append(row)
            if len(rows) >= 2:
                return FactTable(columns=columns, rows=rows[:6])

        if self._looks_comparative(f'{title} {focus}'):
            fallback_table = self._build_label_detail_table(
                title=title,
                focus=focus,
                facts=normalized_facts,
                language_code=language_code,
            )
            if fallback_table is not None:
                return fallback_table
        return None

    def _split_label_detail(self, fact: str) -> tuple[str, str] | None:
        clean = self._fit_text(fact, 170)
        if not clean:
            return None
        separators = [': ', ' — ', ' – ', ' - ', '; ']
        for separator in separators:
            if separator in clean:
                left, right = clean.split(separator, 1)
                left = self._clean_text_artifacts(left)
                right = self._clean_text_artifacts(right)
                if 1 <= len(left.split()) <= 5 and len(right.split()) >= 3:
                    return self._fit_text(left, 26), self._fit_text(right, 72)
        if ', ' in clean:
            left, right = clean.split(', ', 1)
            left = self._clean_text_artifacts(left)
            right = self._clean_text_artifacts(right)
            if 1 <= len(left.split()) <= 5 and len(right.split()) >= 4:
                return self._fit_text(left, 26), self._fit_text(right, 72)
        return None

    def _build_label_detail_table(self, *, title: str, focus: str, facts: list[str], language_code: str) -> FactTable | None:
        pack = self._language_pack(language_code)
        rows: list[list[str]] = []
        for fact in facts[:6]:
            pair = self._split_label_detail(fact)
            if pair is not None:
                rows.append([pair[0], pair[1]])
        if len(rows) < 3:
            return None
        left_header, right_header = pack['comparison_headers']
        columns = [self._fit_text(left_header, 24), self._fit_text(right_header, 26)]
        return FactTable(columns=columns, rows=rows[:6])

    def _fallback_key_terms(self, *, topic_title: str, existing: list[str], needed: int, language_code: str = 'uz') -> list[str]:
        pack = self._language_pack(language_code)
        raw_terms = [*self._topic_keywords(topic_title), *pack['key_terms_seed']]
        result: list[str] = []
        used = set(existing)
        for term in raw_terms:
            text = self._fit_text(term, 60)
            if text and text not in used:
                used.add(text)
                result.append(text)
            if len(result) >= needed:
                break
        return result

    def _fallback_agenda_items(self, *, topic: str, existing: list[str], needed: int, language_code: str = 'uz') -> list[str]:
        pack = self._language_pack(language_code)
        result: list[str] = []
        used = set(existing)
        for item in pack['agenda_seed']:
            text = self._fit_text(item, 70)
            if text not in used:
                used.add(text)
                result.append(text)
            if len(result) >= needed:
                break
        if len(result) < needed:
            fallback = self._fit_text(str(pack['agenda_suffix']).format(topic=topic), 70)
            while len(result) < needed:
                result.append(fallback)
        return result

    def _fallback_summary_points(self, *, topic: str, needed: int, language_code: str = 'uz') -> list[str]:
        pack = self._language_pack(language_code)
        base = self._fit_text(topic, 80) or str(pack['unknown_topic'])
        pool = [self._fit_text(template.format(topic=base), 170) for template in pack['summary_seed']]
        return pool[:needed]

