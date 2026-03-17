from __future__ import annotations

import re
import tempfile
from pathlib import Path
from typing import Any

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.enum.text import MSO_VERTICAL_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

from app.schemas.presentation_plan import PresentationPlan, TopicSection
from app.services.gemini_planner import GeminiPresentationPlanner


class PptxGenerationService:
    def __init__(
        self,
        output_dir: str | None = None,
        gemini_planner: GeminiPresentationPlanner | None = None,
    ) -> None:
        base_dir = Path(output_dir) if output_dir else Path(tempfile.gettempdir()) / 'slide_generator_outputs'
        base_dir.mkdir(parents=True, exist_ok=True)
        self.output_dir = base_dir
        self.gemini_planner = gemini_planner

    def generate(self, payload: dict) -> str:
        plan = self.build_plan(payload)
        return self.render(payload, plan)

    def build_plan(self, payload: dict) -> PresentationPlan:
        topic, presenter_name, slide_count, language_code, pack = self._extract_core(payload)
        if self.gemini_planner and self.gemini_planner.enabled:
            return self.gemini_planner.build_plan(
                topic=topic,
                presenter_name=presenter_name,
                slide_count=slide_count,
                language_code=language_code,
            )
        return self._build_fallback_plan(topic=topic, slide_count=slide_count, pack=pack)

    def render(self, payload: dict, plan: PresentationPlan) -> str:
        topic, presenter_name, _, language_code, pack = self._extract_core(payload)
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        total_slides = 3 + len(plan.sections)

        self._add_title_slide(
            prs,
            presentation_title=plan.presentation_title,
            subtitle=plan.title_subtitle,
            presenter_name=presenter_name,
            agenda_preview=plan.agenda_items[:4],
            page_number=1,
            total_slides=total_slides,
            pack=pack,
        )
        self._add_agenda_slide(
            prs,
            agenda_items=plan.agenda_items,
            agenda_notes=[section.focus for section in plan.sections[:4]],
            presenter_name=presenter_name,
            page_number=2,
            total_slides=total_slides,
            pack=pack,
        )

        current_page = 3
        for section in plan.sections:
            if section.content_type == 'table' and section.table is not None:
                self._add_table_slide(
                    prs,
                    section=section,
                    presenter_name=presenter_name,
                    page_number=current_page,
                    total_slides=total_slides,
                    pack=pack,
                )
            elif section.content_type == 'process':
                self._add_process_slide(
                    prs,
                    section=section,
                    presenter_name=presenter_name,
                    page_number=current_page,
                    total_slides=total_slides,
                    pack=pack,
                )
            else:
                self._add_facts_slide(
                    prs,
                    section=section,
                    presenter_name=presenter_name,
                    page_number=current_page,
                    total_slides=total_slides,
                    pack=pack,
                )
            current_page += 1

        self._add_summary_slide(
            prs,
            title=str(pack['summary']),
            summary_points=plan.summary_points,
            presenter_name=presenter_name,
            page_number=total_slides,
            total_slides=total_slides,
            pack=pack,
        )

        safe_stem = self._safe_filename(f"{topic[:60]}_{presenter_name[:30]}_{language_code}")
        file_path = self.output_dir / f'{safe_stem}.pptx'
        prs.save(file_path)
        return str(file_path)

    @staticmethod
    def _safe_filename(value: str) -> str:
        cleaned = re.sub(r'[^A-Za-z0-9А-Яа-яЎўҚқҒғҲҳ_\- ]+', '', value).strip()
        cleaned = re.sub(r'\s+', '_', cleaned)
        return cleaned or 'presentation'

    def _extract_core(self, payload: dict) -> tuple[str, str, int, str, dict]:
        language_code = str(payload.get('language_code', 'uz') or 'uz')
        pack = self._language_pack(language_code)
        topic = str(payload.get('topic', 'Untitled presentation')).strip()
        presenter_name = str(payload.get('presenter_name', 'Unknown')).strip()
        slide_count = int(payload.get('slide_count', 6) or 6)
        slide_count = min(max(slide_count, 6), 15)
        return topic, presenter_name, slide_count, language_code, pack


    @staticmethod
    def _language_pack(language_code: str) -> dict[str, str | list[str]]:
        if language_code == 'ru':
            return {
                'language_code': 'ru',
                'prepared_by': 'Подготовил',
                'agenda': 'Основные разделы',
                'agenda_note_title': 'Краткий обзор',
                'agenda_subtitle': 'Основные этапы и направления темы.',
                'summary': 'Итоговые выводы',
                'summary_subtitle': 'Ключевые факты и выводы по теме.',
                'key_focus': 'Ключевые темы',
                'focus_label': 'Фокус раздела',
                'cover_points': ['Истоки', 'Развитие', 'Ключевые факты', 'Значение'],
            }
        if language_code == 'en':
            return {
                'language_code': 'en',
                'prepared_by': 'Prepared by',
                'agenda': 'Main sections',
                'agenda_note_title': 'Quick overview',
                'agenda_subtitle': 'The main sections and thematic directions.',
                'summary': 'Final conclusions',
                'summary_subtitle': 'The key facts and concluding takeaways.',
                'key_focus': 'Key themes',
                'focus_label': 'Section focus',
                'cover_points': ['Origins', 'Development', 'Key facts', 'Significance'],
            }
        return {
            'language_code': 'uz',
            'prepared_by': 'Tayyorlagan',
            'agenda': 'Asosiy bo‘limlar',
            'agenda_note_title': 'Qisqacha reja',
            'agenda_subtitle': 'Asosiy davrlar va yo‘nalishlar.',
            'summary': 'Yakuniy xulosalar',
            'summary_subtitle': 'Mavzu bo‘yicha yakuniy faktlar va xulosalar.',
            'key_focus': 'Asosiy yo‘nalishlar',
            'focus_label': 'Bo‘lim markazi',
            'cover_points': ['Boshlanish', 'Rivojlanish', 'Muhim faktlar', 'Ahamiyati'],
        }

    def _build_fallback_plan(self, *, topic: str, slide_count: int, pack: dict) -> PresentationPlan:
        language_code = str(pack.get('language_code', 'uz'))
        if language_code == 'en':
            subtitle = f'{topic}: the main ideas, key facts, and practical significance.'
            sections = [
                TopicSection(content_type='facts', title='Core concepts', focus=f'This section explains the main ideas behind {topic}.', facts=[
                    f'{topic} can be understood through several connected dimensions.',
                    'The key terms and examples help define the topic more clearly.',
                    'Causes, features, and outcomes make the topic easier to compare and explain.',
                    'A structured overview helps show why the topic matters in practice or theory.',
                ]),
                TopicSection(content_type='process', title='Development or sequence', focus=f'This section follows the main stages or sequence related to {topic}.', facts=[
                    'The initial context or starting conditions are identified first.',
                    'The main stages or changes are then described in logical order.',
                    'Important factors that influence the development are highlighted.',
                    'The final state or outcome is summarized clearly.',
                ]),
                TopicSection(content_type='facts', title='Key implications', focus=f'This section highlights why {topic} is important.', facts=[
                    'Examples and comparisons help clarify the meaning of the topic.',
                    'The topic often becomes clearer when viewed from both theory and practice.',
                    'Well-chosen facts make the content easier to remember and explain.',
                    'A concise summary strengthens the overall understanding of the subject.',
                ]),
            ]
            summary = [
                f'{topic} is best understood through its core ideas, examples, and implications.',
                'A clear sequence of facts helps connect the topic’s causes, features, and outcomes.',
                'Carefully selected examples make the subject more memorable and easier to explain.',
            ]
        elif language_code == 'ru':
            subtitle = f'{topic}: основные идеи, ключевые факты и практическое значение темы.'
            sections = [
                TopicSection(content_type='facts', title='Основные понятия', focus=f'Раздел раскрывает ключевые идеи темы {topic}.', facts=[
                    f'{topic} раскрывается через несколько взаимосвязанных аспектов.',
                    'Ключевые термины и примеры помогают точнее понять содержание темы.',
                    'Причины, признаки и последствия делают тему более понятной для сравнения.',
                    'Структурированный обзор помогает показать значимость темы на практике и в теории.',
                ]),
                TopicSection(content_type='process', title='Последовательность развития', focus=f'Раздел показывает основные этапы или последовательность, связанные с {topic}.', facts=[
                    'Сначала обозначается исходный контекст и стартовые условия.',
                    'Затем в логическом порядке раскрываются основные этапы или изменения.',
                    'Подчеркиваются важные факторы, влияющие на развитие темы.',
                    'В завершение кратко показывается итоговое состояние или результат.',
                ]),
                TopicSection(content_type='facts', title='Ключевое значение', focus=f'Раздел показывает, почему {topic} имеет значение.', facts=[
                    'Примеры и сопоставления помогают точнее раскрыть содержание темы.',
                    'Тема становится понятнее при рассмотрении с теоретической и практической сторон.',
                    'Хорошо подобранные факты делают содержание более запоминающимся.',
                    'Краткие выводы помогают закрепить целостное понимание темы.',
                ]),
            ]
            summary = [
                f'{topic} лучше всего раскрывается через основные идеи, примеры и значение.',
                'Последовательное изложение фактов помогает связать причины, признаки и результаты темы.',
                'Удачно подобранные примеры делают содержание более понятным и запоминающимся.',
            ]
        else:
            subtitle = f'{topic} mavzusining asosiy g‘oyalari, muhim faktlari va amaliy ahamiyati.'
            sections = [
                TopicSection(content_type='facts', title='Asosiy tushunchalar', focus=f'Bu bo‘lim {topic} mavzusining tayanch g‘oyalarini yoritadi.', facts=[
                    f'{topic} bir nechta o‘zaro bog‘liq jihatlar orqali tushuniladi.',
                    'Asosiy atamalar va misollar mavzu mazmunini aniqroq ochib beradi.',
                    'Sabab, belgi va natijalarni ajratish mavzuni solishtirishni yengillashtiradi.',
                    'Tartibli sharh mavzuning nazariy va amaliy ahamiyatini ko‘rsatadi.',
                ]),
                TopicSection(content_type='process', title='Rivojlanish ketma-ketligi', focus=f'Bu bo‘lim {topic} bilan bog‘liq asosiy bosqich yoki ketma-ketlikni ko‘rsatadi.', facts=[
                    'Avval boshlang‘ich sharoit va kontekst aniqlanadi.',
                    'So‘ng asosiy bosqichlar yoki o‘zgarishlar mantiqiy tartibda ko‘rsatiladi.',
                    'Rivojlanishga ta’sir qiluvchi muhim omillar ajratib ko‘rsatiladi.',
                    'Yakunida umumiy holat yoki natija qisqacha jamlanadi.',
                ]),
                TopicSection(content_type='facts', title='Muhim xulosalar', focus=f'Bu bo‘lim {topic} mavzusining ahamiyatini ko‘rsatadi.', facts=[
                    'Misollar va taqqoslashlar mavzuning mazmunini ravshanlashtiradi.',
                    'Mavzu nazariy va amaliy nuqtai nazardan ko‘rilganda yaxshiroq tushuniladi.',
                    'To‘g‘ri tanlangan faktlar mavzuni esda qolarli qiladi.',
                    'Qisqa xulosalar mavzu bo‘yicha yaxlit tasavvur beradi.',
                ]),
            ]
            summary = [
                f'{topic} mavzusi asosiy g‘oyalar, misollar va ahamiyat orqali yaxshiroq tushuniladi.',
                'Faktlarni ketma-ket ko‘rish mavzuning sabab, belgi va natijalarini bog‘laydi.',
                'To‘g‘ri tanlangan misollar mazmunni aniqroq va esda qolarli qiladi.',
            ]

        section_count = max(3, slide_count - 3)
        while len(sections) < section_count:
            sections.append(sections[len(sections) % 3])
        agenda_items = [section.title for section in sections[: min(8, len(sections))]]
        return PresentationPlan(
            presentation_title=topic,
            title_subtitle=subtitle,
            agenda_items=agenda_items,
            sections=sections[:section_count],
            summary_points=summary[:5],
        )

    def _topic_profile(self, topic: str) -> dict:
        lower = topic.lower()
        if any(marker in lower for marker in ('oʻzbekiston tarixi', "o'zbekiston tarixi", 'uzbekiston tarixi', 'history of uzbekistan', 'ozbekiston tarixi')):
            return {
                'subtitle': 'Qadimgi davrlardan mustaqillik davrigacha bo‘lgan siyosiy, madaniy va ijtimoiy taraqqiyot yo‘li.',
                'sections': [
                    TopicSection(content_type='facts', title='Qadimgi va ilk davlatlar', focus='Hududdagi ilk sivilizatsiyalar va davlat birlashmalarining shakllanishi.', facts=[
                        'Amudaryo va Sirdaryo oralig‘i qadimdan dehqonchilik, hunarmandchilik va savdo uchun qulay hudud bo‘lgan.',
                        'Baxtariya, So‘g‘d va Xorazm kabi qadimgi birliklar Markaziy Osiyo tarixida muhim siyosiy markazlar bo‘lib xizmat qilgan.',
                        'Hudud Buyuk Ipak yo‘li orqali Sharq va G‘arb madaniyatlari tutashgan maydonga aylangan.',
                        'Arxeologik topilmalar shaharlashuv va yozma madaniyatning juda erta shakllanganini ko‘rsatadi.',
                    ]),
                    TopicSection(content_type='process', title='Islom davri va uyg‘onish bosqichi', focus='VIII–XII asrlarda ilm-fan, shaharlar va madaniyatning yuksalish jarayoni.', facts=[
                        'Arablar kirib kelishi bilan islom dini hududda keng tarqala boshladi.',
                        'Mahalliy sulolalar shakllanib, Buxoro va Samarqand ilmiy markaz sifatida kuchaydi.',
                        'Madrasalar, kutubxonalar va ilmiy muhit rivojlanib, ko‘plab olimlar yetishib chiqdi.',
                        'Hudud musulmon Sharqidagi yirik ilm-fan va madaniyat markazlaridan biriga aylandi.',
                    ]),
                    TopicSection(content_type='facts', title='Temuriylar davri', focus='Amir Temur va temuriylar zamonida davlat boshqaruvi, bunyodkorlik va ilmning yuksalishi.', facts=[
                        'Amir Temur markazlashgan davlat tuzib, Movarounnahr siyosiy qudratini tikladi.',
                        'Samarqand yirik siyosiy va me’moriy markazga aylantirildi.',
                        'Ulug‘bek davrida astronomiya, matematika va madaniy hayot yangi bosqichga ko‘tarildi.',
                        'Registon, observatoriya va ko‘plab me’moriy obidalar bu davr merosini ifodalaydi.',
                    ]),
                    TopicSection(content_type='table', title='Tarixiy davrlar taqqoslanishi', focus='Asosiy davrlar bo‘yicha siyosiy, madaniy va ilmiy xususiyatlarni solishtirish.', facts=[
                        'Har bir davrda davlat boshqaruvi va madaniy rivojlanish darajasi turlicha bo‘lgan.',
                        'Ilm-fan va me’morchilikning yuksalishi ayniqsa uyg‘onish va temuriylar davrida kuchli ko‘ringan.',
                    ], table={
                        'columns': ['Davr', 'Asosiy markaz', 'Ajralib turgan jihat'],
                        'rows': [
                            ['Qadimgi davr', 'Xorazm, So‘g‘d', 'Ilk davlatchilik va savdo yo‘llari'],
                            ['IX–XII asrlar', 'Buxoro, Samarqand', 'Ilm-fan va madaniy uyg‘onish'],
                            ['Temuriylar', 'Samarqand, Hirot', 'Bunyodkorlik va ilmiy taraqqiyot'],
                            ['Mustaqillik davri', 'Toshkent', 'Milliy tiklanish va merosni asrash'],
                        ],
                    }),
                    TopicSection(content_type='facts', title='Xonliklar va yangi davr', focus='Buxoro, Xiva va Qo‘qon xonliklari davrida siyosiy raqobat va mahalliy boshqaruv.', facts=[
                        'Hudud bir necha xonliklarga bo‘linib, markazlashuv zaiflashdi.',
                        'Buxoro, Xiva va Qo‘qon o‘ziga xos siyosiy va iqtisodiy tizimlarga ega bo‘lgan.',
                        'Savdo, hunarmandchilik va qishloq xo‘jaligi ichki iqtisodiyotning asosini tashkil etgan.',
                        'Ichki nizolar tashqi bosim oldida hududning zaiflashishiga sabab bo‘lgan.',
                    ]),
                    TopicSection(content_type='facts', title='Mustamlaka va jadidchilik', focus='Rossiya imperiyasi davri, islohot g‘oyalari va milliy uyg‘onishning kuchayishi.', facts=[
                        'XIX asr oxiriga kelib hudud Rossiya imperiyasi ta’siri ostiga tushdi.',
                        'Mustamlaka boshqaruvi iqtisodiy va siyosiy hayotda keskin o‘zgarishlar olib keldi.',
                        'Jadidlar yangi usul maktablari, matbuot va ma’rifat orqali jamiyatni isloh qilishga intildi.',
                        'Milliy o‘zlik, ta’lim va zamonaviy fikr jadidchilik harakatining asosiy mavzulari bo‘ldi.',
                    ]),
                    TopicSection(content_type='facts', title='Mustaqillik va madaniy meros', focus='1991-yildan keyingi davrda tarixiy xotira, milliy tiklanish va merosni asrash ishlari.', facts=[
                        '1991-yilda O‘zbekiston mustaqillikka erishib, yangi davlat taraqqiyoti yo‘lini boshladi.',
                        'Tarixiy obidalarni restavratsiya qilish va milliy merosni tiklash davlat siyosatining muhim qismiga aylandi.',
                        'Samarqand, Buxoro va Xiva kabi shaharlarga xalqaro e’tibor kuchaydi.',
                        'Tarix va madaniyat milliy o‘zlikni mustahkamlovchi asosiy omillardan biri bo‘lib qoldi.',
                    ]),
                ],
                'summary_points': [
                    'O‘zbekiston tarixi qadimgi sivilizatsiyalar, yirik davlatlar va boy madaniy meros uzviyligiga tayangan.',
                    'Hudud ilm-fan, savdo va me’morchilik markazi sifatida Markaziy Osiyoda alohida o‘rin tutgan.',
                    'Temuriylar davri va jadidchilik harakati tarixiy rivojlanishda burilish nuqtalari bo‘lib xizmat qilgan.',
                    'Mustaqillik davrida tarixiy xotira va madaniy merosni asrash yangi bosqichga ko‘tarildi.',
                ],
            }

        base_facts = [
            f'{topic} mavzusi bir nechta muhim yo‘nalishlar orqali tushuntiriladi.',
            'Asosiy tushunchalar va faktlarni tarixiy yoki amaliy kontekst bilan ko‘rib chiqish mazmunni boyitadi.',
            'Mavzudagi sabab va natijalarni ajratish tushunishni ancha yengillashtiradi.',
            'Misollar va taqqoslashlar mavzuning ahamiyatini yanada ravshanlashtiradi.',
        ]
        return {
            'subtitle': f'{topic} mavzusining asosiy yo‘nalishlari, muhim faktlari va amaliy ahamiyati.',
            'sections': [
                TopicSection(content_type='facts', title='Asosiy mazmun', focus=f'{topic} mavzusining markaziy g‘oyalari va tayanch faktlari.', facts=base_facts),
                TopicSection(content_type='facts', title='Muhim omillar', focus=f'{topic}ga ta’sir qiluvchi sabablar va sharoitlar.', facts=base_facts),
                TopicSection(content_type='process', title='Rivojlanish ketma-ketligi', focus=f'{topic} bo‘yicha asosiy bosqichlar yoki jarayonlar.', facts=[
                    'Boshlang‘ich sharoit va kontekstni aniqlash.',
                    'Asosiy o‘zgarishlar yoki rivojlanish nuqtalarini ko‘rsatish.',
                    'Natijaga ta’sir qilgan muhim omillarni ajratish.',
                    'Bugungi holat yoki yakuniy natijani baholash.',
                ]),
                TopicSection(content_type='table', title='Qisqa taqqoslash', focus=f'{topic} bo‘yicha asosiy jihatlarni jadval orqali jamlash.', facts=[
                    'Jadval mavzudagi asosiy jihatlarni bir joyga to‘playdi.',
                    'Taqqoslash orqali farq va o‘xshashliklar tezroq ko‘rinadi.',
                ], table={
                    'columns': ['Jihat', 'Tavsif', 'Ahamiyati'],
                    'rows': [
                        ['Mazmun', topic, 'Asosiy yo‘nalish'],
                        ['Asosiy omil', 'Tayanch faktlar', 'Tushunishni kuchaytiradi'],
                        ['Natija', 'Xulosa va ta’sir', 'Amaliy ahamiyat'],
                    ],
                }),
            ],
            'summary_points': [
                f'{topic} mavzusi mazmunan bir nechta o‘zaro bog‘liq yo‘nalishlardan tashkil topadi.',
                'Asosiy faktlar va jarayonlarni ketma-ket ko‘rish umumiy tasavvurni kuchaytiradi.',
                'Taqqoslash va misollar mavzuni esda qolarli va aniq qiladi.',
            ],
        }

    @staticmethod
    def _normalize_text(value: Any, max_chars: int | None = None) -> str:
        text = '' if value is None else str(value)
        text = text.replace(' ', ' ')
        text = text.replace('\"', '"').replace("\'", "'")
        text = text.replace('“', '"').replace('”', '"').replace('‘', "'").replace('’', "'")
        text = re.sub(r'\s+', ' ', text).strip()
        if (text.startswith('{') and text.endswith('}')) or (text.startswith('[') and text.endswith(']')):
            text = text[1:-1].strip()
        text = text.replace('{', '').replace('}', '').replace('[', '').replace(']', '')
        text = re.sub(r'\s*([,:;])\s*', r'\1 ', text)
        text = re.sub(r'\s+', ' ', text).strip(" \"'\n\t-–—")
        if max_chars and len(text) > max_chars:
            text = text[:max_chars].rstrip(' ,;:-.')
        return text


    @staticmethod
    def _fit_frame(text_frame, *, max_size: int | float, min_size: int | float = 10, bold: bool = False) -> None:
        text_frame.word_wrap = True
        try:
            text_frame.fit_text(max_size=max_size, font_family='Arial', bold=bold)
        except Exception:
            pass
        for paragraph in text_frame.paragraphs:
            for run in paragraph.runs:
                current = run.font.size.pt if run.font.size else float(max_size)
                if current < float(min_size):
                    run.font.size = Pt(float(min_size))
                elif current > float(max_size):
                    run.font.size = Pt(float(max_size))

    @staticmethod
    def _content_bounds(has_subtitle: bool) -> tuple[float, float]:
        return (2.08 if has_subtitle else 1.86, 6.58)

    @staticmethod
    def _estimate_lines(text: str, width_inches: float, font_size: float, *, bullet: bool = False) -> int:
        cleaned = re.sub(r'\s+', ' ', str(text or '')).strip()
        if not cleaned:
            return 1
        words = cleaned.split()
        capacity = max(
            12,
            int(
                width_inches
                * (
                    21 if font_size <= 11 else 19 if font_size <= 12 else 17 if font_size <= 13.5 else 15 if font_size <= 15 else 13
                )
            ),
        )
        current = 2 if bullet else 0
        lines = 1
        for word in words:
            token_len = len(word) + 1
            if current + token_len > capacity:
                lines += 1
                current = len(word)
            else:
                current += token_len
        return max(1, lines)

    def _estimate_bullet_block_height(
        self,
        items: list[str],
        *,
        width_inches: float,
        font_size: float,
        space_after_pt: float,
    ) -> float:
        line_height_inches = (font_size * 1.18) / 72
        gap_inches = space_after_pt / 72
        total = 0.0
        for item in items:
            lines = self._estimate_lines(item, width_inches, font_size, bullet=True)
            total += (lines * line_height_inches) + gap_inches
        return total

    def _balance_items_for_columns(self, items: list[str], *, width_inches: float, font_size: float) -> list[list[str]]:
        if len(items) <= 3:
            return [items, []]
        best_split: list[list[str]] | None = None
        best_score: float | None = None
        for split in range(1, len(items)):
            left = items[:split]
            right = items[split:]
            left_h = self._estimate_bullet_block_height(left, width_inches=width_inches, font_size=font_size, space_after_pt=8)
            right_h = self._estimate_bullet_block_height(right, width_inches=width_inches, font_size=font_size, space_after_pt=8)
            score = abs(left_h - right_h) + (abs(len(left) - len(right)) * 0.12)
            if best_score is None or score < best_score:
                best_score = score
                best_split = [left, right]
        return best_split or [items, []]

    def _select_facts_layout(self, items: list[str], *, panel_height: float) -> dict:
        presets = [
            {'columns': 1, 'font_size': 18.0, 'space_after': 12, 'width': 10.35},
            {'columns': 1, 'font_size': 16.8, 'space_after': 11, 'width': 10.35},
            {'columns': 1, 'font_size': 15.5, 'space_after': 10, 'width': 10.35},
            {'columns': 2, 'font_size': 15.0, 'space_after': 10, 'width': 4.92},
            {'columns': 2, 'font_size': 14.0, 'space_after': 9, 'width': 4.92},
            {'columns': 2, 'font_size': 13.0, 'space_after': 8, 'width': 4.92},
            {'columns': 2, 'font_size': 12.0, 'space_after': 7, 'width': 4.92},
        ]
        best: dict | None = None
        best_score: float | None = None
        for preset in presets:
            if preset['columns'] == 1:
                columns = [items]
                required_h = self._estimate_bullet_block_height(items, width_inches=preset['width'], font_size=preset['font_size'], space_after_pt=preset['space_after'])
            else:
                columns = self._balance_items_for_columns(items, width_inches=preset['width'], font_size=preset['font_size'])
                non_empty = [col for col in columns if col]
                required_h = max(
                    self._estimate_bullet_block_height(col, width_inches=preset['width'], font_size=preset['font_size'], space_after_pt=preset['space_after'])
                    for col in non_empty
                )
            fill_ratio = required_h / max(panel_height, 0.1)
            if required_h <= panel_height:
                score = abs(0.84 - fill_ratio)
                if preset['columns'] == 2 and len(items) <= 3:
                    score += 0.15
                if best_score is None or score < best_score:
                    best_score = score
                    best = {**preset, 'columns_data': columns, 'required_h': required_h}
        if best:
            return best
        fallback = presets[-1]
        columns = self._balance_items_for_columns(items, width_inches=fallback['width'], font_size=fallback['font_size'])
        return {**fallback, 'columns_data': columns, 'required_h': panel_height}

    def _write_bullet_block(
        self,
        frame,
        items: list[str],
        *,
        font_size: float,
        color: RGBColor,
        space_after_pt: float,
        align=PP_ALIGN.LEFT,
        min_size: int = 10,
    ) -> None:
        frame.clear()
        frame.word_wrap = True
        frame.vertical_anchor = MSO_VERTICAL_ANCHOR.TOP
        frame.margin_left = Inches(0.02)
        frame.margin_right = Inches(0.02)
        frame.margin_top = Inches(0.01)
        frame.margin_bottom = Inches(0.01)
        for idx, item in enumerate(items):
            p = frame.paragraphs[0] if idx == 0 else frame.add_paragraph()
            p.alignment = align
            p.text = f'• {self._normalize_text(item)}'
            p.level = 0
            p.space_after = Pt(space_after_pt)
            for run in p.runs:
                run.font.size = Pt(font_size)
                run.font.color.rgb = color
        self._fit_frame(frame, max_size=int(round(font_size)), min_size=min_size)

    def _set_cell_text(
        self,
        cell,
        text: str,
        *,
        font_size: float,
        bold: bool = False,
        color: tuple[int, int, int] = (31, 41, 55),
        min_size: float = 9.5,
    ) -> None:
        cell.text = self._normalize_text(text)
        frame = cell.text_frame
        frame.word_wrap = True
        frame.margin_left = Inches(0.03)
        frame.margin_right = Inches(0.03)
        frame.margin_top = Inches(0.03)
        frame.margin_bottom = Inches(0.03)
        for p in frame.paragraphs:
            for run in p.runs:
                run.font.size = Pt(font_size)
                run.font.bold = bold
                run.font.color.rgb = RGBColor(*color)
        self._fit_frame(frame, max_size=font_size, min_size=min_size, bold=bold)


    @staticmethod
    def _facts_variant(items: list[str]) -> str:
        if not items:
            return 'bullets'
        max_len = max(len(item) for item in items)
        avg_len = sum(len(item) for item in items) / max(1, len(items))
        if len(items) in {3, 4} and max_len <= 95 and avg_len <= 78:
            return 'cards'
        if len(items) in {3, 4} and avg_len <= 118:
            return 'spotlight'
        return 'bullets'

    @staticmethod
    def _fact_card_coords(item_count: int) -> list[tuple[float, float, float, float]]:
        if item_count == 3:
            return [
                (0.92, 2.42, 3.72, 2.58),
                (4.81, 2.42, 3.72, 2.58),
                (8.70, 2.42, 3.72, 2.58),
            ]
        if item_count == 4:
            return [
                (0.92, 2.20, 5.12, 1.82),
                (7.30, 2.20, 5.12, 1.82),
                (0.92, 4.18, 5.12, 1.82),
                (7.30, 4.18, 5.12, 1.82),
            ]
        if item_count == 2:
            return [
                (1.00, 2.58, 5.40, 2.24),
                (6.93, 2.58, 5.40, 2.24),
            ]
        return []

    def _render_fact_cards(self, slide, *, items: list[str]) -> None:
        coords = self._fact_card_coords(len(items))
        colors = [
            RGBColor(239, 246, 255),
            RGBColor(238, 242, 255),
            RGBColor(236, 253, 245),
            RGBColor(255, 247, 237),
        ]
        font_size = 14.4 if max((len(item) for item in items), default=0) <= 72 else 13.2
        for index, item in enumerate(items):
            x, y, w, h = coords[index]
            shape = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
            shape.fill.solid()
            shape.fill.fore_color.rgb = RGBColor(255, 255, 255)
            shape.line.color.rgb = RGBColor(191, 219, 254)

            stripe = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(0.18))
            stripe.fill.solid()
            stripe.fill.fore_color.rgb = colors[index % len(colors)]
            stripe.line.fill.background()

            badge = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.OVAL, Inches(x + 0.18), Inches(y + 0.26), Inches(0.46), Inches(0.46))
            badge.fill.solid()
            badge.fill.fore_color.rgb = RGBColor(30, 64, 175)
            badge.line.fill.background()
            badge_frame = badge.text_frame
            badge_frame.clear()
            badge_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
            badge_p = badge_frame.paragraphs[0]
            badge_p.alignment = PP_ALIGN.CENTER
            badge_run = badge_p.add_run()
            badge_run.text = str(index + 1)
            badge_run.font.size = Pt(12)
            badge_run.font.bold = True
            badge_run.font.color.rgb = RGBColor(255, 255, 255)

            box = slide.shapes.add_textbox(Inches(x + 0.38), Inches(y + 0.86), Inches(w - 0.60), Inches(h - 1.00))
            frame = box.text_frame
            frame.clear()
            frame.word_wrap = True
            frame.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
            frame.margin_left = Inches(0.04)
            frame.margin_right = Inches(0.04)
            p = frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            run = p.add_run()
            run.text = item
            run.font.size = Pt(font_size)
            run.font.color.rgb = RGBColor(30, 41, 59)
            self._fit_frame(frame, max_size=font_size, min_size=10.8)

    def _render_focus_spotlight(self, slide, *, focus: str, items: list[str], pack: dict) -> None:
        left = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, Inches(0.92), Inches(2.18), Inches(4.12), Inches(3.98))
        left.fill.solid()
        left.fill.fore_color.rgb = RGBColor(239, 246, 255)
        left.line.color.rgb = RGBColor(191, 219, 254)

        frame = left.text_frame
        frame.clear()
        frame.word_wrap = True
        frame.margin_left = Inches(0.18)
        frame.margin_right = Inches(0.18)
        frame.margin_top = Inches(0.14)
        header = frame.paragraphs[0]
        header.alignment = PP_ALIGN.LEFT
        run = header.add_run()
        run.text = str(pack['focus_label'])
        run.font.size = Pt(13.2)
        run.font.bold = True
        run.font.color.rgb = RGBColor(30, 64, 175)

        focus_p = frame.add_paragraph()
        focus_p.alignment = PP_ALIGN.LEFT
        focus_p.space_before = Pt(8)
        focus_p.space_after = Pt(10)
        focus_run = focus_p.add_run()
        focus_run.text = self._normalize_text(focus)
        focus_run.font.size = Pt(15)
        focus_run.font.color.rgb = RGBColor(30, 41, 59)
        self._fit_frame(frame, max_size=15, min_size=11.2)

        right = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, Inches(5.36), Inches(2.18), Inches(7.00), Inches(3.98))
        right.fill.solid()
        right.fill.fore_color.rgb = RGBColor(255, 255, 255)
        right.line.color.rgb = RGBColor(203, 213, 225)

        inner = slide.shapes.add_textbox(Inches(5.68), Inches(2.46), Inches(6.34), Inches(3.42))
        self._write_bullet_block(
            inner.text_frame,
            items,
            font_size=13.6,
            color=RGBColor(31, 41, 55),
            space_after_pt=9,
            min_size=10.2,
        )

    @staticmethod
    def _process_coords(item_count: int) -> list[tuple[float, float, float, float]]:
        if item_count <= 3:
            width = 3.55 if item_count == 3 else 4.45
            gap = 0.38
            total_width = (width * item_count) + (gap * max(0, item_count - 1))
            start_x = max(0.78, (13.333 - total_width) / 2)
            return [(start_x + i * (width + gap), 2.42, width, 1.92) for i in range(item_count)]
        if item_count == 4:
            return [
                (0.92, 2.12, 5.15, 1.78),
                (7.26, 2.12, 5.15, 1.78),
                (0.92, 4.18, 5.15, 1.78),
                (7.26, 4.18, 5.15, 1.78),
            ]
        return [
            (0.58, 2.02, 4.02, 1.62),
            (4.66, 2.02, 4.02, 1.62),
            (8.74, 2.02, 4.02, 1.62),
            (2.62, 4.12, 4.02, 1.62),
            (6.70, 4.12, 4.02, 1.62),
        ]

    @staticmethod
    def _agenda_layout(item_count: int) -> dict[str, Any]:
        if item_count <= 5:
            return {'columns': 1, 'left': 0.90, 'top': 2.04, 'col_width': 5.72, 'item_height': 0.58, 'gap': 0.12}
        if item_count <= 7:
            return {'columns': 1, 'left': 0.88, 'top': 2.00, 'col_width': 5.74, 'item_height': 0.50, 'gap': 0.09}
        return {'columns': 2, 'left': 0.82, 'top': 2.00, 'col_width': 2.76, 'item_height': 0.52, 'gap': 0.10, 'col_gap': 0.16}

    def _base_slide(
        self,
        prs: Presentation,
        *,
        title: str,
        presenter_name: str,
        page_number: int,
        total_slides: int,
        pack: dict,
        subtitle: str | None = None,
        background_rgb: tuple[int, int, int] = (248, 250, 252),
    ):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        background = slide.background.fill
        background.solid()
        background.fore_color.rgb = RGBColor(*background_rgb)

        top_band = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.62))
        top_band.fill.solid()
        top_band.fill.fore_color.rgb = RGBColor(30, 64, 175)
        top_band.line.fill.background()

        accent = slide.shapes.add_shape(
            MSO_AUTO_SHAPE_TYPE.RECTANGLE,
            Inches(0.72),
            Inches(0.82),
            Inches(0.10),
            Inches(1.22 if subtitle else 0.82),
        )
        accent.fill.solid()
        accent.fill.fore_color.rgb = RGBColor(96, 165, 250)
        accent.line.fill.background()

        if title:
            title_box = slide.shapes.add_textbox(Inches(0.96), Inches(0.80), Inches(11.1), Inches(0.92))
            title_frame = title_box.text_frame
            title_frame.clear()
            title_frame.word_wrap = True
            title_frame.margin_left = 0
            title_frame.margin_right = 0
            p = title_frame.paragraphs[0]
            run = p.add_run()
            run.text = self._normalize_text(title)
            run.font.size = Pt(25)
            run.font.bold = True
            run.font.color.rgb = RGBColor(15, 23, 42)
            self._fit_frame(title_frame, max_size=25, min_size=18, bold=True)

        if subtitle:
            subtitle_box = slide.shapes.add_textbox(Inches(0.96), Inches(1.48), Inches(11.25), Inches(0.68))
            subtitle_frame = subtitle_box.text_frame
            subtitle_frame.clear()
            subtitle_frame.word_wrap = True
            subtitle_frame.margin_left = 0
            subtitle_frame.margin_right = 0
            p = subtitle_frame.paragraphs[0]
            run = p.add_run()
            run.text = self._normalize_text(subtitle)
            run.font.size = Pt(12.5)
            run.font.color.rgb = RGBColor(71, 85, 105)
            self._fit_frame(subtitle_frame, max_size=12, min_size=10)

        footer_line = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, Inches(0.72), Inches(6.84), Inches(11.86), Inches(0.02))
        footer_line.fill.solid()
        footer_line.fill.fore_color.rgb = RGBColor(203, 213, 225)
        footer_line.line.fill.background()

        author_box = slide.shapes.add_textbox(Inches(0.76), Inches(6.90), Inches(5.8), Inches(0.26))
        author_frame = author_box.text_frame
        author_frame.clear()
        author_run = author_frame.paragraphs[0].add_run()
        author_run.text = f"{pack['prepared_by']}: {self._normalize_text(presenter_name, max_chars=40)}"
        author_run.font.size = Pt(10)
        author_run.font.color.rgb = RGBColor(100, 116, 139)

        page_box = slide.shapes.add_textbox(Inches(11.45), Inches(6.90), Inches(1.05), Inches(0.26))
        page_frame = page_box.text_frame
        page_frame.clear()
        page_frame.paragraphs[0].alignment = PP_ALIGN.RIGHT
        page_run = page_frame.paragraphs[0].add_run()
        page_run.text = f'{page_number}/{total_slides}'
        page_run.font.size = Pt(10)
        page_run.font.color.rgb = RGBColor(100, 116, 139)
        return slide

    def _add_title_slide(
        self,
        prs: Presentation,
        *,
        presentation_title: str,
        subtitle: str,
        presenter_name: str,
        agenda_preview: list[str],
        page_number: int,
        total_slides: int,
        pack: dict,
    ) -> None:
        slide = self._base_slide(
            prs,
            title='',
            presenter_name=presenter_name,
            page_number=page_number,
            total_slides=total_slides,
            pack=pack,
            background_rgb=(239, 246, 255),
        )

        title_box = slide.shapes.add_textbox(Inches(0.92), Inches(1.28), Inches(6.6), Inches(1.68))
        frame = title_box.text_frame
        frame.clear()
        frame.word_wrap = True
        p = frame.paragraphs[0]
        run = p.add_run()
        run.text = self._normalize_text(presentation_title)
        run.font.size = Pt(28)
        run.font.bold = True
        run.font.color.rgb = RGBColor(15, 23, 42)
        self._fit_frame(frame, max_size=28, min_size=20, bold=True)

        subtitle_box = slide.shapes.add_textbox(Inches(0.95), Inches(2.92), Inches(6.2), Inches(1.18))
        subtitle_frame = subtitle_box.text_frame
        subtitle_frame.clear()
        subtitle_frame.word_wrap = True
        subtitle_run = subtitle_frame.paragraphs[0].add_run()
        subtitle_run.text = self._normalize_text(subtitle)
        subtitle_run.font.size = Pt(15)
        subtitle_run.font.color.rgb = RGBColor(51, 65, 85)
        self._fit_frame(subtitle_frame, max_size=15, min_size=11)

        card = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, Inches(7.7), Inches(1.56), Inches(4.58), Inches(3.78))
        card.fill.solid()
        card.fill.fore_color.rgb = RGBColor(255, 255, 255)
        card.line.color.rgb = RGBColor(191, 219, 254)
        card_frame = card.text_frame
        card_frame.clear()
        card_frame.word_wrap = True
        card_frame.margin_left = Inches(0.18)
        card_frame.margin_right = Inches(0.18)
        card_frame.margin_top = Inches(0.12)
        card_frame.margin_bottom = Inches(0.12)

        header = card_frame.paragraphs[0]
        header.alignment = PP_ALIGN.CENTER
        run = header.add_run()
        run.text = str(pack['key_focus'])
        run.font.size = Pt(17)
        run.font.bold = True
        run.font.color.rgb = RGBColor(30, 64, 175)

        for item in (agenda_preview or list(pack['cover_points'])[:4])[:4]:
            paragraph = card_frame.add_paragraph()
            paragraph.text = f'• {self._normalize_text(item)}'
            paragraph.space_after = Pt(9)
            for run in paragraph.runs:
                run.font.size = Pt(15)
                run.font.color.rgb = RGBColor(30, 41, 59)
        self._fit_frame(card_frame, max_size=17, min_size=11)

    def _add_agenda_slide(self, prs: Presentation, *, agenda_items: list[str], agenda_notes: list[str], presenter_name: str, page_number: int, total_slides: int, pack: dict) -> None:
        slide = self._base_slide(
            prs,
            title=str(pack['agenda']),
            presenter_name=presenter_name,
            page_number=page_number,
            total_slides=total_slides,
            pack=pack,
            subtitle=str(pack['agenda_subtitle']),
        )
        visible_items = [self._normalize_text(item) for item in agenda_items[:8] if self._normalize_text(item)]
        layout = self._agenda_layout(len(visible_items))
        note_left = 7.08
        note_width = 5.1
        body_font = 13.8 if layout['columns'] == 1 and len(visible_items) <= 5 else 12.4 if layout['columns'] == 1 else 11.7

        for index, item in enumerate(visible_items, start=1):
            if layout['columns'] == 1:
                x = layout['left']
                y = layout['top'] + (layout['item_height'] + layout['gap']) * (index - 1)
            else:
                left_count = (len(visible_items) + 1) // 2
                col = 0 if index <= left_count else 1
                row = (index - 1) if col == 0 else (index - 1 - left_count)
                x = layout['left'] + col * (layout['col_width'] + layout['col_gap'])
                y = layout['top'] + row * (layout['item_height'] + layout['gap'])
                note_left = 6.82
                note_width = 5.34

            shape = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(layout['col_width']), Inches(layout['item_height']))
            shape.fill.solid()
            shape.fill.fore_color.rgb = RGBColor(255, 255, 255)
            shape.line.color.rgb = RGBColor(191, 219, 254)
            frame = shape.text_frame
            frame.clear()
            frame.word_wrap = True
            frame.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
            frame.margin_left = Inches(0.12)
            frame.margin_right = Inches(0.12)
            run = frame.paragraphs[0].add_run()
            run.text = f'{index}. {item}'
            run.font.size = Pt(body_font)
            run.font.color.rgb = RGBColor(15, 23, 42)
            self._fit_frame(frame, max_size=body_font, min_size=10.8)

        note = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, Inches(note_left), Inches(2.10), Inches(note_width), Inches(3.62))
        note.fill.solid()
        note.fill.fore_color.rgb = RGBColor(219, 234, 254)
        note.line.fill.background()
        note_frame = note.text_frame
        note_frame.clear()
        note_frame.word_wrap = True
        note_frame.margin_left = Inches(0.16)
        note_frame.margin_right = Inches(0.16)
        note_frame.margin_top = Inches(0.10)
        note_frame.margin_bottom = Inches(0.10)
        title_paragraph = note_frame.paragraphs[0]
        title_paragraph.alignment = PP_ALIGN.CENTER
        run = title_paragraph.add_run()
        run.text = str(pack['agenda_note_title'])
        run.font.size = Pt(15.5)
        run.font.bold = True
        run.font.color.rgb = RGBColor(30, 64, 175)
        for item in (agenda_notes or visible_items)[:4]:
            p = note_frame.add_paragraph()
            p.text = f'• {self._normalize_text(item)}'
            p.space_after = Pt(7)
            for run in p.runs:
                run.font.size = Pt(12.0)
                run.font.color.rgb = RGBColor(30, 41, 59)
        self._fit_frame(note_frame, max_size=15, min_size=10.5)


    def _add_facts_slide(self, prs: Presentation, *, section: TopicSection, presenter_name: str, page_number: int, total_slides: int, pack: dict) -> None:
        slide = self._base_slide(
            prs,
            title=section.title,
            presenter_name=presenter_name,
            page_number=page_number,
            total_slides=total_slides,
            pack=pack,
            subtitle=section.focus,
        )

        facts = [self._normalize_text(item) for item in section.facts[:6] if self._normalize_text(item)]
        if not facts:
            return

        variant = self._facts_variant(facts)
        if variant == 'cards':
            self._render_fact_cards(slide, items=facts[:4])
            return
        if variant == 'spotlight':
            self._render_focus_spotlight(slide, focus=section.focus, items=facts[:4], pack=pack)
            return

        content_top, content_bottom = self._content_bounds(bool(section.focus))
        panel_left = 0.78
        panel_top = content_top
        panel_width = 11.78
        panel_height = content_bottom - content_top

        panel = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, Inches(panel_left), Inches(panel_top), Inches(panel_width), Inches(panel_height))
        panel.fill.solid()
        panel.fill.fore_color.rgb = RGBColor(255, 255, 255)
        panel.line.color.rgb = RGBColor(203, 213, 225)

        accent = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, Inches(panel_left + 0.18), Inches(panel_top + 0.22), Inches(0.10), Inches(panel_height - 0.44))
        accent.fill.solid()
        accent.fill.fore_color.rgb = RGBColor(191, 219, 254)
        accent.line.fill.background()

        layout = self._select_facts_layout(facts, panel_height=panel_height - 0.48)
        if layout['columns'] == 1:
            text_box = slide.shapes.add_textbox(Inches(panel_left + 0.42), Inches(panel_top + 0.22), Inches(10.92), Inches(panel_height - 0.38))
            self._write_bullet_block(
                text_box.text_frame,
                layout['columns_data'][0],
                font_size=layout['font_size'],
                color=RGBColor(31, 41, 55),
                space_after_pt=layout['space_after'],
                min_size=10.2,
            )
            return

        left_box = slide.shapes.add_textbox(Inches(panel_left + 0.40), Inches(panel_top + 0.22), Inches(5.02), Inches(panel_height - 0.36))
        right_box = slide.shapes.add_textbox(Inches(panel_left + 5.98), Inches(panel_top + 0.22), Inches(5.02), Inches(panel_height - 0.36))

        self._write_bullet_block(left_box.text_frame, layout['columns_data'][0], font_size=layout['font_size'], color=RGBColor(31, 41, 55), space_after_pt=layout['space_after'], min_size=10.2)
        self._write_bullet_block(right_box.text_frame, layout['columns_data'][1], font_size=layout['font_size'], color=RGBColor(31, 41, 55), space_after_pt=layout['space_after'], min_size=10.2)

        divider = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, Inches(panel_left + 5.56), Inches(panel_top + 0.26), Inches(0.02), Inches(panel_height - 0.52))
        divider.fill.solid()
        divider.fill.fore_color.rgb = RGBColor(226, 232, 240)
        divider.line.fill.background()

    def _add_process_slide(self, prs: Presentation, *, section: TopicSection, presenter_name: str, page_number: int, total_slides: int, pack: dict) -> None:
        slide = self._base_slide(
            prs,
            title=section.title,
            presenter_name=presenter_name,
            page_number=page_number,
            total_slides=total_slides,
            pack=pack,
            subtitle=section.focus,
        )

        items = [self._normalize_text(item) for item in section.facts[:5] if self._normalize_text(item)]
        if not items:
            return
        coords = self._process_coords(len(items))
        colors = [
            RGBColor(219, 234, 254),
            RGBColor(224, 231, 255),
            RGBColor(220, 252, 231),
            RGBColor(254, 243, 199),
            RGBColor(254, 226, 226),
        ]
        max_len = max((len(item) for item in items), default=0)
        body_size = 12.5 if max_len <= 70 else 11.5 if max_len <= 100 else 10.5

        for idx, item in enumerate(items):
            x, y, w, h = coords[idx]
            shape = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
            shape.fill.solid()
            shape.fill.fore_color.rgb = colors[idx % len(colors)]
            shape.line.color.rgb = RGBColor(191, 219, 254)

            badge = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.OVAL, Inches(x + 0.16), Inches(y + 0.16), Inches(0.54), Inches(0.54))
            badge.fill.solid()
            badge.fill.fore_color.rgb = RGBColor(30, 64, 175)
            badge.line.fill.background()
            badge_frame = badge.text_frame
            badge_frame.clear()
            badge_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
            badge_p = badge_frame.paragraphs[0]
            badge_p.alignment = PP_ALIGN.CENTER
            badge_run = badge_p.add_run()
            badge_run.text = str(idx + 1)
            badge_run.font.size = Pt(14)
            badge_run.font.bold = True
            badge_run.font.color.rgb = RGBColor(255, 255, 255)

            text_box = slide.shapes.add_textbox(Inches(x + 0.28), Inches(y + 0.82), Inches(w - 0.56), Inches(h - 0.98))
            frame = text_box.text_frame
            frame.clear()
            frame.word_wrap = True
            frame.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
            p = frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            run = p.add_run()
            run.text = item
            run.font.size = Pt(body_size)
            run.font.color.rgb = RGBColor(51, 65, 85)
            self._fit_frame(frame, max_size=int(round(body_size)), min_size=9)

        if len(items) == 5:
            arrow = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.CHEVRON, Inches(5.64), Inches(3.34), Inches(2.05), Inches(0.48))
            arrow.fill.solid()
            arrow.fill.fore_color.rgb = RGBColor(191, 219, 254)
            arrow.line.fill.background()


    def _add_table_slide(self, prs: Presentation, *, section: TopicSection, presenter_name: str, page_number: int, total_slides: int, pack: dict) -> None:
        slide = self._base_slide(
            prs,
            title=section.title,
            presenter_name=presenter_name,
            page_number=page_number,
            total_slides=total_slides,
            pack=pack,
            subtitle=section.focus,
        )
        assert section.table is not None

        content_top, content_bottom = self._content_bounds(bool(section.focus))
        area_height = content_bottom - content_top

        rows = len(section.table.rows) + 1
        cols = len(section.table.columns)
        max_cell_len = max((len(self._normalize_text(cell)) for row in section.table.rows for cell in row), default=20)

        use_side_note = cols >= 3
        if use_side_note:
            table_left = 0.76
            table_top = content_top
            table_width = 9.35
            note_left = 10.32
            note_width = 2.06
        else:
            table_left = 0.78
            table_top = content_top + 0.06
            table_width = 11.78
            note_left = note_width = 0.0

        table_shape = slide.shapes.add_table(rows, cols, Inches(table_left), Inches(table_top), Inches(table_width), Inches(area_height - (0.06 if not use_side_note else 0.0)))
        table = table_shape.table

        def col_weight(index: int) -> float:
            header = self._normalize_text(section.table.columns[index])
            header_weight = max(8, len(header) * 1.9)
            row_weight = max((len(self._normalize_text(row[index])) for row in section.table.rows if index < len(row)), default=12)
            return max(header_weight, row_weight)

        weights = [col_weight(i) for i in range(cols)]
        total_weight = sum(weights) or cols
        min_width = 2.15 if cols == 2 else 1.35
        for idx in range(cols):
            ratio = weights[idx] / total_weight
            width = max(min_width, round(table_width * ratio, 2))
            table.columns[idx].width = Inches(width)

        header_height = 0.56 if cols <= 3 else 0.54
        body_height = max(0.62 if rows <= 4 else 0.54, (area_height - header_height) / max(1, rows - 1))
        table.rows[0].height = Inches(header_height)
        for row_idx in range(1, rows):
            table.rows[row_idx].height = Inches(body_height)

        if cols == 2:
            header_font = 12.4
            body_font = 12.0 if max_cell_len <= 54 else 11.2
            body_min = 10.4
        else:
            header_font = 11.5 if cols <= 3 else 10.8
            body_font = 11.1 if rows <= 4 and max_cell_len <= 48 else 10.2 if max_cell_len <= 70 else 9.6
            body_min = 8.8

        for col_idx, column_name in enumerate(section.table.columns):
            cell = table.cell(0, col_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(219, 234, 254)
            self._set_cell_text(cell, column_name, font_size=header_font, bold=True, color=(30, 64, 175), min_size=max(10.0, header_font - 1.2))

        for row_idx, row in enumerate(section.table.rows, start=1):
            for col_idx, value in enumerate(row):
                cell = table.cell(row_idx, col_idx)
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(255, 255, 255) if row_idx % 2 else RGBColor(248, 250, 252)
                self._set_cell_text(cell, value, font_size=body_font, bold=False, color=(31, 41, 55), min_size=body_min)

        if not use_side_note:
            return

        note = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, Inches(note_left), Inches(content_top), Inches(note_width), Inches(area_height))
        note.fill.solid()
        note.fill.fore_color.rgb = RGBColor(239, 246, 255)
        note.line.color.rgb = RGBColor(191, 219, 254)

        note_header_box = slide.shapes.add_textbox(Inches(note_left + 0.10), Inches(content_top + 0.16), Inches(note_width - 0.20), Inches(0.30))
        note_header_frame = note_header_box.text_frame
        note_header_frame.clear()
        note_header_frame.word_wrap = True
        note_header_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        run = note_header_frame.paragraphs[0].add_run()
        run.text = str(pack['focus_label'])
        run.font.size = Pt(12.4)
        run.font.bold = True
        run.font.color.rgb = RGBColor(30, 64, 175)

        note_body_box = slide.shapes.add_textbox(Inches(note_left + 0.12), Inches(content_top + 0.52), Inches(note_width - 0.24), Inches(area_height - 0.68))
        note_frame = note_body_box.text_frame
        note_frame.clear()
        note_frame.word_wrap = True
        focus_p = note_frame.paragraphs[0]
        focus_p.text = self._normalize_text(section.focus)
        focus_p.space_after = Pt(7)
        for run in focus_p.runs:
            run.font.size = Pt(10.8)
            run.font.color.rgb = RGBColor(51, 65, 85)

        side_facts = section.facts[:1] if rows >= 5 else section.facts[:2]
        for fact in side_facts:
            p = note_frame.add_paragraph()
            p.text = f'• {self._normalize_text(fact)}'
            p.space_after = Pt(4)
            for run in p.runs:
                run.font.size = Pt(9.8)
                run.font.color.rgb = RGBColor(31, 41, 55)
        self._fit_frame(note_frame, max_size=11.2, min_size=9.0)

    def _add_summary_slide(self, prs: Presentation, *, title: str, summary_points: list[str], presenter_name: str, page_number: int, total_slides: int, pack: dict) -> None:
        slide = self._base_slide(
            prs,
            title=title,
            presenter_name=presenter_name,
            page_number=page_number,
            total_slides=total_slides,
            pack=pack,
            subtitle=str(pack['summary_subtitle']),
        )

        items = [self._normalize_text(item) for item in summary_points[:5] if self._normalize_text(item)]
        if not items:
            return

        if len(items) == 4:
            coords = [
                (0.90, 2.18, 5.40, 1.92),
                (7.03, 2.18, 5.40, 1.92),
                (0.90, 4.28, 5.40, 1.92),
                (7.03, 4.28, 5.40, 1.92),
            ]
        elif len(items) == 5:
            coords = [
                (0.82, 2.10, 3.88, 1.82),
                (4.73, 2.10, 3.88, 1.82),
                (8.64, 2.10, 3.88, 1.82),
                (2.78, 4.12, 3.88, 1.82),
                (6.69, 4.12, 3.88, 1.82),
            ]
        elif len(items) == 3:
            coords = [
                (0.92, 2.55, 3.72, 2.48),
                (4.80, 2.55, 3.72, 2.48),
                (8.68, 2.55, 3.72, 2.48),
            ]
        else:
            coords = [
                (1.18, 2.60, 5.20, 2.34),
                (6.95, 2.60, 5.20, 2.34),
            ]

        max_len = max((len(item) for item in items), default=0)
        font_size = 15.2 if max_len <= 75 else 14.0 if max_len <= 110 else 12.8

        for idx, item in enumerate(items):
            x, y, w, h = coords[idx]
            shape = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
            shape.fill.solid()
            shape.fill.fore_color.rgb = RGBColor(255, 255, 255)
            shape.line.color.rgb = RGBColor(191, 219, 254)

            stripe = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(0.16))
            stripe.fill.solid()
            stripe.fill.fore_color.rgb = RGBColor(219, 234, 254)
            stripe.line.fill.background()

            frame = shape.text_frame
            frame.clear()
            frame.word_wrap = True
            frame.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
            frame.margin_left = Inches(0.16)
            frame.margin_right = Inches(0.16)
            frame.margin_top = Inches(0.10)
            frame.margin_bottom = Inches(0.10)
            p = frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            run = p.add_run()
            run.text = item
            run.font.size = Pt(font_size)
            run.font.color.rgb = RGBColor(30, 41, 59)
            self._fit_frame(frame, max_size=font_size, min_size=12.2)
