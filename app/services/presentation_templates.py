from __future__ import annotations

import json
from pathlib import Path
from typing import Any


class PresentationTemplateRegistry:
    DEFAULT_TEMPLATE_ID = 'aurora_clean'

    def __init__(self, templates_dir: str | Path | None = None) -> None:
        base_dir = Path(__file__).resolve().parents[1] / 'presentation_templates'
        self.templates_dir = Path(templates_dir) if templates_dir else base_dir
        self._templates = self._load_templates()

    def _load_templates(self) -> dict[str, dict[str, Any]]:
        templates: dict[str, dict[str, Any]] = {}
        if not self.templates_dir.exists():
            return templates

        for path in sorted(self.templates_dir.glob('*.json')):
            payload = json.loads(path.read_text(encoding='utf-8'))
            template_id = str(payload.get('id') or path.stem).strip()
            if not template_id:
                continue
            payload['id'] = template_id
            preview_path = str(payload.get('preview_path') or '').strip()
            if preview_path:
                payload['preview_path'] = str((self.templates_dir / preview_path).resolve())
            templates[template_id] = payload
        return templates

    def get(self, template_id: str | None) -> dict[str, Any]:
        if template_id and template_id in self._templates:
            return self._templates[template_id]
        if self.DEFAULT_TEMPLATE_ID in self._templates:
            return self._templates[self.DEFAULT_TEMPLATE_ID]
        if self._templates:
            return next(iter(self._templates.values()))
        return self._fallback_template()

    def list_templates(self) -> list[dict[str, Any]]:
        return sorted(
            self._templates.values(),
            key=lambda item: (
                int(item.get('sort_order', 10_000) or 10_000),
                str(item.get('name') or item.get('id') or ''),
            ),
        )

    @property
    def preview_collage_path(self) -> Path:
        return self.templates_dir / 'previews' / 'template-demo-collage.png'

    @staticmethod
    def _fallback_template() -> dict[str, Any]:
        return {
            'id': 'aurora_clean',
            'name': 'Aurora Clean',
            'button_label': 'Aurora Clean',
            'theme': {
                'font_family': 'Aptos',
                'background': '#f8fafc',
                'cover_background': '#eff6ff',
                'primary': '#1e40af',
                'secondary': '#60a5fa',
                'accent': '#bfdbfe',
                'text': '#0f172a',
                'muted': '#475569',
                'footer': '#64748b',
                'surface': '#ffffff',
                'surface_alt': '#f8fafc',
                'border': '#cbd5e1',
                'on_primary': '#ffffff',
                'table_header_bg': '#dbeafe',
                'table_header_text': '#1e40af',
                'process_colors': ['#dbeafe', '#e0e7ff', '#dcfce7', '#fef3c7', '#fee2e2'],
            },
            'layout': {
                'header_style': 'top_band',
                'cover_style': 'split_card',
                'agenda_style': 'list_note',
                'facts_style': 'auto',
                'process_style': 'cards',
                'summary_style': 'cards',
                'accent_style': 'academic_grid',
            },
            'prompt_guidance': [
                'Use concise academic structure with balanced facts, process sections and one useful comparison table when it is natural.',
                'Avoid generic wording and keep every fact slide-ready.',
            ],
        }
