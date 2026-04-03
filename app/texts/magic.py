from __future__ import annotations

from html import escape
from typing import Any


def format_money(amount: int | float | None) -> str:
    return f"{int(amount or 0):,}".replace(',', ' ')


def magic_hook_text(context: dict[str, Any]) -> str:
    price = format_money(context.get('price_uzs'))
    maintenance_note = ''
    if context.get('maintenance_enabled'):
        maintenance_note = (
            "\n\n⚠️ <b>Hozir premium bo‘lim maintenance rejimida.</b> "
            "Hisobni ko‘rib chiqishingiz mumkin, lekin yangi buyurtmalar vaqtincha ochilmagan."
        )

    return (
        "<b>✨ Magic Slayd</b>\n\n"
        "Oddiy taqdimot emas, <b>premium darajadagi</b> vizual ta’sir beradigan, kuchli struktura va tayyor template asosida yig‘iladigan slaydlar shu yerda tayyorlanadi.\n\n"
        "🔥 Himoya, pitch va executive format uchun kuchliroq chiqish\n"
        "🎯 Har bir pack premium template va aniq oqim bilan ishlaydi\n"
        f"💸 1 ta Magic Slayd narxi: <b>{price} so‘m</b>\n"
        "⏱ Mablag‘ faqat PPTX muvaffaqiyatli yuborilgandan keyin yechiladi"
        f"{maintenance_note}"
    )


def magic_account_text(context: dict[str, Any]) -> str:
    balance = format_money(context.get('balance_uzs'))
    price = format_money(context.get('price_uzs'))
    available_presentations = int(context.get('available_presentations', 0) or 0)
    status_line = (
        "✅ <b>Balans kamida 1 ta premium taqdimot uchun yetarli.</b>"
        if context.get('can_afford')
        else "⚠️ <b>Balans hozircha premium taqdimot yaratish uchun yetarli emas.</b>"
    )

    details = [
        "<b>💼 Magic Slayd hisobingiz</b>",
        '',
        f"• Joriy balans: <b>{balance} so‘m</b>",
        f"• 1 ta Magic Slayd narxi: <b>{price} so‘m</b>",
        f"• Mavjud imkoniyat: <b>{available_presentations}</b> ta premium taqdimot",
        '',
        status_line,
    ]

    if context.get('maintenance_enabled'):
        details.extend(['', "🛠 Premium yaratish bo‘limi vaqtincha maintenance rejimida."])

    if not context.get('webapp_configured'):
        details.extend(['', "⚙️ Yaratish oynasi hali ulanmagan, shuning uchun yaratishni boshlash tugmasi vaqtincha ishlamaydi."])

    return '\n'.join(details)


def magic_topup_text(cards: list[dict[str, Any]]) -> str:
    lines = [
        "<b>💳 Hisobni to‘ldirish</b>",
        '',
        "Quyidagi kartalardan biriga to‘lov qilishingiz mumkin. So‘ng mos summani tanlab, chekni yuborasiz.",
        '',
        "<u>Faol qabul kartalari</u>",
    ]
    if not cards:
        lines.append("Hozircha faol qabul kartalari mavjud emas.")
    else:
        for index, card in enumerate(cards, start=1):
            card_number = str(card.get('full_number') or card.get('masked_number') or '—')
            lines.append(
                f"{index}. <code>{escape(card_number)}</code> — <b>{escape(str(card.get('card_holder') or '—'))}</b>"
            )
    lines.extend(['', "Pastdagi summalardan birini tanlang."])
    return '\n'.join(lines)


def magic_topup_unavailable_text() -> str:
    return (
        "<b>💳 Hisobni to‘ldirish</b>\n\n"
        "Hozircha faol qabul kartalari qo‘shilmagan. Administrator kartalarni kiritgach, shu bo‘lim orqali balansni to‘ldirishingiz mumkin bo‘ladi."
    )


def magic_receipt_prompt_text(amount_uzs: int, cards: list[dict[str, Any]]) -> str:
    amount = format_money(amount_uzs)
    lines = [
        "<b>🧾 To‘lov cheki kutilmoqda</b>",
        '',
        f"Quyidagi kartalardan biriga <b>{amount} so‘m</b> miqdorida to‘lov qiling va uni tasdiqlovchi chekni yuboring.",
        "Chek <b>PDF</b> yoki <b>screenshot/rasm</b> bo‘lishi mumkin.",
        '',
        "<u>Qabul kartalari</u>",
    ]
    for index, card in enumerate(cards, start=1):
        card_number = str(card.get('full_number') or card.get('masked_number') or '—')
        lines.append(
            f"{index}. <code>{escape(card_number)}</code> — <b>{escape(str(card.get('card_holder') or '—'))}</b>"
        )
    lines.extend(['', "Chek yuborilgach, so‘rov moderatorga jo‘natiladi."])
    return '\n'.join(lines)


def magic_receipt_received_text(amount_uzs: int) -> str:
    amount = format_money(amount_uzs)
    return (
        "<b>✅ To‘lov cheki qabul qilindi</b>\n\n"
        f"• So‘rov summasi: <b>{amount} so‘m</b>\n"
        "• Holat: <b>moderatsiyada</b>\n\n"
        "To‘lov odatda <b>24 soat ichida</b> ko‘rib chiqiladi. Natija shu suhbatda yuboriladi.\n\n"
        "Noqulayliklar uchun uzr."
    )


def magic_start_ready_text(context: dict[str, Any]) -> str:
    balance = format_money(context.get('balance_uzs'))
    price = format_money(context.get('price_uzs'))
    return (
        "<b>🚀 Magic Slayd yaratishni boshlash</b>\n\n"
        f"• Joriy balans: <b>{balance} so‘m</b>\n"
        f"• 1 ta premium taqdimot narxi: <b>{price} so‘m</b>\n\n"
        "Quyidagi tugmani bosish orqali mavzu, til va kerakli ma’lumotlarni kiriting.\n"
        "<b>Muhim:</b> mablag‘ faqat tayyor PPTX sizga muvaffaqiyatli yuborilgandan keyin yechiladi."
    )


def magic_start_insufficient_text(context: dict[str, Any]) -> str:
    balance = format_money(context.get('balance_uzs'))
    price = format_money(context.get('price_uzs'))
    missing = max(0, int(context.get('price_uzs', 0) or 0) - int(context.get('balance_uzs', 0) or 0))
    return (
        "<b>⚠️ Balans yetarli emas</b>\n\n"
        f"• Joriy balans: <b>{balance} so‘m</b>\n"
        f"• Kerakli summa: <b>{price} so‘m</b>\n"
        f"• Yetishmayotgani: <b>{format_money(missing)} so‘m</b>\n\n"
        "Premium taqdimot yaratish uchun avval hisobingizni to‘ldiring."
    )


def magic_webapp_not_ready_text() -> str:
    return (
        "<b>⚙️ Yaratish oynasi hozircha tayyor emas</b>\n\n"
        "Mavzu va boshqa ma’lumotlarni kiritish oynasi hali ulanmagan. Birozdan keyin qayta urinib ko‘ring."
    )


def magic_maintenance_text() -> str:
    return (
        "<b>🛠 Magic Slayd vaqtincha yopiq</b>\n\n"
        "Premium bo‘limda texnik sozlash ishlari ketmoqda. Birozdan keyin qayta urinib ko‘ring."
    )


def magic_webapp_received_text() -> str:
    return (
        "<b>📥 Magic Slayd buyurtmasi qabul qilindi</b>\n\n"
        "Kiritilgan ma’lumotlar saqlandi.\n"
        "Balansingiz hozircha o‘zgarmaydi va mablag‘ faqat tayyor PPTX yuborilgandan keyin yechiladi."
    )


def magic_order_queued_text(template_name: str, ahead_count: int) -> str:
    queue_note = (
        "Sizning buyurtmangiz hozir ishlov berilmoqda."
        if ahead_count <= 0
        else f"Oldingizda <b>{ahead_count}</b> ta premium buyurtma bor."
    )
    return (
        "<b>📥 Magic Slayd buyurtmasi navbatga qo‘shildi</b>\n\n"
        f"• Template: <b>{template_name}</b>\n"
        f"• Holat: <b>navbatda</b>\n"
        f"• Navbat: {queue_note}\n\n"
        "Tayyor PPTX shu suhbatga yuboriladi. Mablag‘ faqat muvaffaqiyatli yuborilgandan keyin yechiladi."
    )


def magic_order_existing_text(template_name: str | None, ahead_count: int) -> str:
    template_note = f"• Joriy buyurtma: <b>{template_name}</b>\n" if template_name else ''
    queue_note = (
        "Buyurtmangiz hozir ishlab turibdi."
        if ahead_count <= 0
        else f"Oldingizda <b>{ahead_count}</b> ta premium buyurtma bor."
    )
    return (
        "<b>⏳ Sizda faol Magic Slayd buyurtmasi bor</b>\n\n"
        f"{template_note}"
        f"• Holat: <b>navbatda / ishlovda</b>\n"
        f"• Navbat: {queue_note}\n\n"
        "Avval shu buyurtma yakunlansin, keyin yangi premium buyurtma yuborishingiz mumkin."
    )


def magic_order_progress_text(payload: dict[str, Any], percent: int, stage_key: str) -> str:
    topic = str((payload.get('variables') or {}).get('topic') or payload.get('template_name') or 'Magic Slayd').strip()
    stage_texts = {
        'queued': 'Navbat tekshirilyapti',
        'analysis': 'Template va kontent tahlil qilinyapti',
        'rendering': 'Premium PPTX yig‘ilyapti',
        'done': 'Tayyor',
    }
    stage = stage_texts.get(stage_key, 'Ishlanmoqda')
    return (
        "<b>✨ Magic Slayd generatsiya jarayoni</b>\n\n"
        f"• Mavzu: <b>{topic}</b>\n"
        f"• Jarayon: <b>{stage}</b>\n"
        f"• Bajarilish: <b>{percent}%</b>"
    )


def magic_order_failed_text() -> str:
    return (
        "<b>❌ Magic Slayd tayyorlab bo‘lmadi</b>\n\n"
        "Premium buyurtma vaqtincha xatolik sabab yakunlanmadi.\n"
        "Balansingiz o‘zgarmadi. Birozdan keyin qayta urinib ko‘ring."
    )


def magic_order_balance_missing_text(price_uzs: int) -> str:
    return (
        "<b>⚠️ Premium buyurtma to‘xtatildi</b>\n\n"
        f"Bu buyurtma uchun kamida <b>{format_money(price_uzs)} so‘m</b> balans kerak edi.\n"
        "Balansingiz hozir yetarli emas, shu sabab buyurtma ishga tushmadi."
    )


def magic_order_success_caption(payload: dict[str, Any], template_name: str, price_uzs: int) -> str:
    topic = str((payload.get('variables') or {}).get('topic') or template_name).strip()
    return (
        "<b>✨ Magic Slayd tayyor</b>\n\n"
        f"• Template: <b>{template_name}</b>\n"
        f"• Mavzu: <b>{topic}</b>\n"
        f"• Narx: <b>{format_money(price_uzs)} so‘m</b>\n\n"
        "Balans yakuniy hisobi pastdagi holat xabarida ko‘rsatiladi."
    )


def magic_order_done_text(template_name: str, price_uzs: int, balance_uzs: int, *, charged: bool) -> str:
    charged_line = (
        f"• Yechilgan summa: <b>{format_money(price_uzs)} so‘m</b>\n• Qolgan balans: <b>{format_money(balance_uzs)} so‘m</b>"
        if charged
        else "• Balansdan yechishda alohida tekshiruv talab bo‘ldi."
    )
    return (
        "<b>✅ Magic Slayd yuborildi</b>\n\n"
        f"• Template: <b>{template_name}</b>\n"
        f"{charged_line}\n\n"
        "Premium PPTX shu chatga yuborildi."
    )


def magic_order_charge_issue_text() -> str:
    return (
        "<b>⚠️ Balans yechimi tekshiruvga tushdi</b>\n\n"
        "PPTX sizga yuborildi, lekin balansni yakuniy hisoblashda texnik holat yuz berdi.\n"
        "Administrator bilan bog‘lanib qo‘ying, biz buni alohida tekshiramiz."
    )


def magic_topup_approved_text(amount_uzs: int, balance_uzs: int) -> str:
    return (
        "<b>✅ To‘lov tasdiqlandi</b>\n\n"
        f"• Qo‘shilgan summa: <b>{format_money(amount_uzs)} so‘m</b>\n"
        f"• Yangi balans: <b>{format_money(balance_uzs)} so‘m</b>\n\n"
        "Endi Magic Slayd premium bo‘limidan foydalanishingiz mumkin."
    )


def magic_topup_rejected_text(amount_uzs: int) -> str:
    return (
        "<b>❌ To‘lov rad etildi</b>\n\n"
        f"• So‘rov summasi: <b>{format_money(amount_uzs)} so‘m</b>\n"
        "• Sabab: <b>to‘lov tasdiqlanmadi</b>\n\n"
        "Agar xatolik bo‘lgan deb hisoblasangiz, chekni qayta yuborib ko‘ring yoki administrator bilan bog‘laning."
    )


def magic_admin_settings_text(context: dict[str, Any]) -> str:
    settings = context['settings']
    maintenance = 'yoqilgan' if settings.get('maintenance_enabled') else 'o‘chirilgan'
    return (
        "<b>✨ Magic Slayd sozlamalari</b>\n\n"
        f"• 1 ta taqdimot narxi: <b>{format_money(settings.get('price_per_presentation'))} so‘m</b>\n"
        f"• Vaqtinchalik yopish holati: <b>{maintenance}</b>\n"
        f"• Faol kartalar soni: <b>{sum(1 for card in context.get('cards', []) if card.get('is_active'))}</b>\n"
        f"• Jami kartalar: <b>{len(context.get('cards', []))}</b>\n"
        f"• Kutilayotgan to‘lovlar: <b>{context.get('pending_count', 0)}</b>\n"
        f"• Yaratish sahifasi: <b>{'ulangan' if context.get('webapp_configured') else 'ulanmagan'}</b>"
    )


def magic_admin_price_prompt_text(current_price: int) -> str:
    return (
        "<b>💰 Magic Slayd narxini yangilash</b>\n\n"
        f"Joriy narx: <b>{format_money(current_price)} so‘m</b>\n\n"
        "Yangi qiymatni butun son ko‘rinishida yuboring.\n"
        "Masalan: <code>15000</code>"
    )


def magic_admin_cards_text(cards: list[dict[str, Any]]) -> str:
    lines = [
        "<b>💳 Magic Slayd kartalari</b>",
        '',
    ]
    if not cards:
        lines.append("Hozircha karta qo‘shilmagan.")
        return '\n'.join(lines)

    for index, card in enumerate(cards, start=1):
        status = 'faol' if card.get('is_active') else 'faolsiz'
        lines.append(
            f"{index}. <code>{escape(str(card.get('masked_number') or '—'))}</code> — <b>{escape(str(card.get('card_holder') or '—'))}</b> — <b>{status}</b>"
        )
    return '\n'.join(lines)


def magic_admin_card_prompt_text() -> str:
    return (
        "<b>➕ Yangi karta qo‘shish</b>\n\n"
        "Kartani quyidagi formatda yuboring:\n"
        "<code>8600 1234 5678 9012 | CARD HOLDER</code>"
    )


def magic_admin_card_text(card: dict[str, Any]) -> str:
    status = 'faol' if card.get('is_active') else 'faolsiz'
    return (
        "<b>💳 Karta tafsiloti</b>\n\n"
        f"• Raqam: <code>{escape(str(card.get('masked_number') or '—'))}</code>\n"
        f"• Egasi: <b>{escape(str(card.get('card_holder') or '—'))}</b>\n"
        f"• Holati: <b>{status}</b>"
    )


def magic_admin_pending_text(topups: list[dict[str, Any]]) -> str:
    lines = [
        "<b>🧾 Kutilayotgan to‘lovlar</b>",
        '',
    ]
    if not topups:
        lines.append("Hozircha moderatsiyada turgan to‘lovlar yo‘q.")
        return '\n'.join(lines)

    for index, topup in enumerate(topups, start=1):
        username = topup.get('username')
        username_text = f"@{str(username).lstrip('@')}" if username else '—'
        lines.append(
            f"{index}. <b>{escape(str(topup.get('full_name') or 'Noma’lum'))}</b> — <b>{format_money(topup.get('amount_uzs'))} so‘m</b> — {username_text}"
        )
    lines.append('')
    lines.append("Chekni qayta ko‘rish uchun pastdagi tugmalardan birini tanlang.")
    return '\n'.join(lines)
