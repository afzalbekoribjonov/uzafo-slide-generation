from __future__ import annotations

import re


class ContentValidator:
    @staticmethod
    def validate_topic(topic: str) -> tuple[bool, str]:
        topic = re.sub(r'\s+', ' ', str(topic or '')).strip()
        if len(topic) < 5:
            return False, "Mavzu biroz qisqa. Iltimos, aniqroq mavzu yozing."
        if len(topic) > 200:
            return False, "Mavzu juda uzun. Iltimos, uni qisqaroq qilib yozing."
        if re.search(r'(.)\1{4,}', topic):
            return False, "Mavzuda bir xil belgi juda ko'p takrorlangan. Iltimos, qayta yozing."
        if re.search(r'https?://|www\.', topic, re.IGNORECASE):
            return False, "Havola emas, taqdimot mavzusini matn ko'rinishida yuboring."
        if re.search(r'[!@#$%^&*_=+<>?/\\|]{4,}', topic):
            return False, "Mavzuda maxsus belgilar ko'p. Iltimos, oddiy matn bilan yozing."
        if len(topic.split()) == 1 and len(topic) < 8:
            return False, "Mavzuni biroz aniqroq yozing. Masalan: 'AI ta'limda' kabi."
        return True, "OK"

    @staticmethod
    def validate_presenter_name(name: str) -> tuple[bool, str]:
        name = re.sub(r'\s+', ' ', str(name or '')).strip()
        if len(name) < 2:
            return False, "Ism juda qisqa. Iltimos, taqdimotda ko'rinadigan ismni yozing."
        if len(name) > 100:
            return False, "Ism juda uzun. Iltimos, qisqaroq yozing."
        special_count = sum(1 for char in name if not char.isalnum() and not char.isspace() and char not in {"'", "-", "."})
        if special_count > max(3, len(name) * 0.3):
            return False, "Ismda maxsus belgilar ko'p. Iltimos, oddiy ism yozing."
        return True, "OK"
