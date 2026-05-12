from __future__ import annotations

from html import escape


_PROGRESS_STAGE_TEXTS = {
    'queued': 'So‘rovingiz qabul qilindi. Navbatdagi o‘rnini kutmoqda.',
    'research': 'Mavzu bo‘yicha kerakli faktlar va misollar saralanmoqda.',
    'planning': 'Slaydlar mazmuni, tartibi va dizayn ko‘rinishi rejalashtirilmoqda.',
    'rendering': 'Sahifalar tanlangan dizayn asosida chiroyli qilib yig‘ilmoqda.',
    'uploading': 'Fayl tayyorlanib, sizga yuborilmoqda.',
    'done': 'Taqdimot tayyor.',
}


def _progress_bar(percent: int) -> str:
    percent = max(0, min(100, int(percent)))
    filled = max(1, round(percent / 10)) if percent > 0 else 0
    return '█' * filled + '░' * (10 - filled)


def _available_generation_label(user: dict, available_generations: int) -> str:
    if user.get('generation_unlimited'):
        return '♾ <b>Cheksiz</b>'
    if user.get('generation_access_blocked'):
        return '⏸ <b>Vaqtincha yopiq</b>'
    return f'🎟 <b>{available_generations}</b>'


def main_menu_text(full_name: str) -> str:
    return (
        f"<b>👋 Assalomu alaykum, {escape(full_name)}!</b>\n\n"
        "Men sizga mavzu asosida <b>tayyor PowerPoint taqdimot</b> yaratib beraman. "
        "Mavzuni yozasiz, dizaynni tanlaysiz, qolgan ishni bot bajaradi.\n\n"
        "Boshlash uchun kerakli bo‘limni tanlang."
    )



def status_text(user: dict, available_generations: int) -> str:
    full_name = escape(user.get('full_name', 'Noma’lum foydalanuvchi'))
    telegram_id = user.get('telegram_id', '-')
    generated_count = int(user.get('generated_count', 0) or 0)
    successful = int(user.get('successful_generations', 0) or 0)
    referral_count = int(user.get('referral_count', 0) or 0)
    referral_credits = int(user.get('referral_credits', 0) or 0)
    bonus_credits = int(user.get('bonus_generation_credits', 0) or 0)

    return (
        "<b>📊 Hisobingiz holati</b>\n\n"
        f"👤 Ism: <b>{full_name}</b>\n"
        f"🆔 Telegram ID: <code>{telegram_id}</code>\n\n"
        f"🎞 Jami so‘rovlar: <b>{generated_count}</b>\n"
        f"✅ Tayyor taqdimotlar: <b>{successful}</b>\n"
        f"👥 Tasdiqlangan takliflar: <b>{referral_count}</b>\n"
        f"🎁 Taklif orqali olingan imkoniyatlar: <b>{referral_credits}</b>\n"
        f"➕ Qo‘shimcha imkoniyatlar: <b>{bonus_credits}</b>\n"
        f"🎟 Hozirgi yaratish imkoniyati: {_available_generation_label(user, available_generations)}"
    )



def invite_text(referral_link: str, available_generations: int, referral_count: int, user: dict) -> str:
    return (
        "<b>👥 Do‘stlarni taklif qilish</b>\n\n"
        "Do‘stingiz botga sizning havolangiz orqali kirib, kerakli obunani tasdiqlasa, sizga <b>1 ta qo‘shimcha taqdimot yaratish imkoniyati</b> qo‘shiladi.\n\n"
        f"✅ Tasdiqlangan takliflar: <b>{referral_count}</b>\n"
        f"🎟 Hozirgi imkoniyat: {_available_generation_label(user, available_generations)}\n\n"
        "<b>🔗 Shaxsiy taklif havolangiz:</b>\n"
        f"<code>{escape(referral_link)}</code>"
    )



def referrals_text(referrals: list[dict]) -> str:
    if not referrals:
        return (
            "<b>👥 Takliflar ro‘yxati</b>\n\n"
            "Hozircha sizning havolangiz orqali kirgan foydalanuvchilar yo‘q."
        )

    lines = ["<b>👥 Takliflar ro‘yxati</b>", ""]
    confirmed = 0

    for index, item in enumerate(referrals, start=1):
        status = '✅ tasdiqlangan' if item.get('counted') else '🕓 kutilmoqda'
        if item.get('counted'):
            confirmed += 1
        lines.append(f"{index}. <code>{item.get('invitee_id')}</code> — <b>{status}</b>")

    lines.append('')
    lines.append(f"✅ Jami tasdiqlangan takliflar: <b>{confirmed}</b>")
    return '\n'.join(lines)



def help_text() -> str:
    return (
        "<b>❓ Yordam</b>\n\n"
        "🎞 <b>Bot nima qiladi?</b>\n"
        "• Mavzu bo‘yicha PowerPoint taqdimot yaratadi\n"
        "• 6 dan 12 tagacha slayd tanlash mumkin\n"
        "• 8 xil dizayn ko‘rinishi bor\n"
        "• Rang, shrift, jadval va matn joylashuvi avtomatik moslanadi\n\n"
        "🚀 <b>Qanday ishlataman?</b>\n"
        "1. <b>Slayd yaratish</b> tugmasini bosing\n"
        "2. Mavzuni aniq yozing\n"
        "3. Tayyorlagan ismni kiriting\n"
        "4. Slayd soni va dizaynni tanlang\n"
        "5. Tilni tanlab, tasdiqlang\n\n"
        "🎁 <b>Eslatma:</b>\n"
        "Birinchi taqdimot bepul. Keyingi imkoniyatlarni do‘st taklif qilish yoki administrator orqali olish mumkin."
    )



def contact_text(support_contact: str) -> str:
    return (
        "<b>☎️ Aloqa</b>\n\n"
        "Savol, taklif yoki yordam kerak bo‘lsa, quyidagi manzilga yozishingiz mumkin:\n\n"
        f"<b>{escape(support_contact)}</b>"
    )



def subscription_text(channels: list[dict]) -> str:
    if not channels:
        return (
            "<b>📢 Obuna</b>\n\n"
            "Hozircha obuna bo‘lish kerak bo‘lgan kanal yo‘q."
        )

    lines = [
        "<b>📢 Davom etish uchun kanallarga obuna bo‘ling</b>",
        '',
        "Quyidagi kanallarga a’zo bo‘ling, keyin <b>✅ Obunani tekshirish</b> tugmasini bosing.",
        '',
    ]

    for index, channel in enumerate(channels, start=1):
        title = escape(channel.get('title') or channel.get('username') or f'Kanal {index}')
        lines.append(f"{index}. {title}")

    return '\n'.join(lines)



def subscription_failed_text(channels: list[dict]) -> str:
    lines = [
        "<b>📢 Obuna hali tasdiqlanmadi</b>",
        '',
        "Quyidagi kanallarga obuna bo‘lganingiz ko‘rinmayapti:",
        '',
    ]

    for index, channel in enumerate(channels, start=1):
        title = escape(channel.get('title') or channel.get('username') or f'Kanal {index}')
        lines.append(f"{index}. {title}")

    lines.append('')
    lines.append("Obuna bo‘lgach, yana <b>✅ Obunani tekshirish</b> tugmasini bosing.")
    return '\n'.join(lines)





def create_generation_blocked_text() -> str:
    return (
        "<b>⏸ Slayd yaratish hozircha yopiq</b>\n\n"
        "Hozirda siz uchun taqdimot yaratish vaqtincha cheklangan. Keyinroq qayta urinib ko‘ring yoki yordam uchun aloqa bo‘limiga yozing."
    )

def create_credit_missing_text() -> str:
    return (
        "<b>🎞 Slayd yaratish</b>\n\n"
        "Hozir hisobingizda yangi taqdimot yaratish imkoniyati qolmagan.\n\n"
        "🎁 Birinchi taqdimot — <b>bepul</b>\n"
        "👥 Keyingi imkoniyatlar — <b>do‘st taklif qilish</b> yoki <b>administrator orqali</b>\n\n"
        "Qo‘shimcha imkoniyat olish uchun do‘stingizni taklif qiling yoki aloqa bo‘limiga yozing."
    )



def create_topic_prompt_text() -> str:
    return (
        "<b>📝 1-bosqich / 6</b>\n\n"
        "Taqdimot mavzusini yozing. Mavzu qanchalik aniq bo‘lsa, slaydlar shunchalik chiroyli va mazmunli chiqadi.\n\n"
        "Masalan: <i>Sun’iy intellekt yordamida shaxsiylashtirilgan ta’lim tizimi</i>"
    )



def create_presenter_prompt_text() -> str:
    return (
        "<b>👤 2-bosqich / 6</b>\n\n"
        "Endi taqdimotda ko‘rinadigan ismni yuboring.\n\n"
        "Masalan: <i>Ali Valiyev</i>"
    )



def create_slide_count_prompt_text() -> str:
    return (
        "<b>📑 3-bosqich / 6</b>\n\n"
        "Nechta slayd kerakligini tanlang.\n\n"
        "📌 Tanlash mumkin: <b>6</b> dan <b>12</b> gacha."
    )


def create_template_prompt_text() -> str:
    return (
        "<b>🎨 4-bosqich / 6</b>\n\n"
        "Taqdimot dizaynini tanlang. Har bir ko‘rinishda ranglar, shriftlar, jadvallar va matn joylashuvi boshqacha."
    )


def create_template_preview_caption() -> str:
    return (
        "<b>🎨 Dizayn tanlash</b>\n\n"
        "Rasmda 8 xil demo ko‘rinish bor. Kerakli uslubni pastdagi tugmalardan tanlang."
    )



def create_language_prompt_text() -> str:
    return (
        "<b>🌐 5-bosqich / 6</b>\n\n"
        "Taqdimot qaysi tilda tayyorlansin?"
    )


def create_pdf_choice_prompt_text() -> str:
    return (
        "<b>📄 6-bosqich / 6</b>\n\n"
        "Taqdimotni PDF shaklida ham yuboraymi?\n\n"
        "✅ <b>Ha</b> — PPTX bilan birga PDF faylni ham olasiz.\n"
        "❌ <b>Yo‘q</b> — Faqat PowerPoint (pptx) fayli yuboriladi."
    )


def create_confirmation_text(data: dict) -> str:
    topic = escape(data.get('topic', '-'))
    presenter_name = escape(data.get('presenter_name', '-'))
    slide_count = data.get('slide_count', '-')
    template_name = escape(data.get('template_name', '-'))
    language_name = escape(data.get('language_name', '-'))
    wants_pdf = 'Ha' if data.get('wants_pdf') else 'Yo‘q'

    return (
        "<b>✅ Hammasi tayyormi?</b>\n\n"
        f"📝 Mavzu: <b>{topic}</b>\n"
        f"👤 Ism: <b>{presenter_name}</b>\n"
        f"📑 Slaydlar: <b>{slide_count}</b>\n"
        f"🎨 Dizayn: <b>{template_name}</b>\n"
        f"🌐 Til: <b>{language_name}</b>\n"
        f"📄 PDF: <b>{wants_pdf}</b>\n\n"
        "Ma’lumotlar to‘g‘ri bo‘lsa, <b>✅ Tasdiqlash</b> tugmasini bosing."
    )



def create_queued_text(data: dict, ahead_count: int) -> str:
    topic = escape(data.get('topic', '-'))
    presenter_name = escape(data.get('presenter_name', '-'))
    slide_count = data.get('slide_count', '-')
    template_name = escape(data.get('template_name', '-'))
    language_name = escape(data.get('language_name', '-'))
    queue_text = (
        'Navbat bo‘sh. Tez orada tayyorlash boshlanadi.'
        if ahead_count == 0
        else f'Sizdan oldinda <b>{ahead_count}</b> ta so‘rov bor.'
    )

    return (
        "<b>📥 So‘rovingiz qabul qilindi</b>\n\n"
        f"📝 Mavzu: <b>{topic}</b>\n"
        f"👤 Ism: <b>{presenter_name}</b>\n"
        f"📑 Slaydlar: <b>{slide_count}</b>\n"
        f"🎨 Dizayn: <b>{template_name}</b>\n"
        f"🌐 Til: <b>{language_name}</b>\n\n"
        f"{queue_text}\n"
        "Jarayon holatini shu yerda ko‘rsatib boraman."
    )



def create_already_queued_text(ahead_count: int) -> str:
    queue_text = (
        'Taqdimotingiz hozir tayyorlanmoqda.'
        if ahead_count == 0
        else f'Sizdan oldinda <b>{ahead_count}</b> ta so‘rov bor.'
    )
    return (
        "<b>⏳ Taqdimot tayyorlanmoqda</b>\n\n"
        "Sizda allaqachon bitta faol so‘rov bor. U tugaguncha yangisini boshlash shart emas.\n\n"
        f"{queue_text}"
    )



def create_generation_progress_text(data: dict, percent: int, stage_key: str) -> str:
    topic = escape(data.get('topic', '-'))
    stage_text = _PROGRESS_STAGE_TEXTS.get(stage_key, 'Jarayon davom etmoqda.')
    bar = _progress_bar(percent)
    return (
        "<b>⚙️ Taqdimot tayyorlanmoqda</b>\n\n"
        f"📝 Mavzu: <b>{topic}</b>\n"
        f"📌 Holat: {stage_text}\n\n"
        f"<code>{bar}</code> <b>{percent}%</b>"
    )



def create_generation_success_caption(data: dict, bot_username: str | None = None) -> str:
    topic = escape(data.get('topic', '-'))
    slide_count = data.get('slide_count', '-')
    template_name = escape(data.get('template_name', '-'))
    clean_username = str(bot_username or '').strip().lstrip('@')
    prepared_by = (
        f"🤖 Taqdimot @{escape(clean_username)} orqali tayyorlandi!"
        if clean_username
        else "🤖 Taqdimot bot orqali tayyorlandi!"
    )
    return (
        "<b>✅ Taqdimot tayyor!</b>\n\n"
        f"📝 Mavzu: <b>{topic}</b>\n"
        f"📑 Slaydlar: <b>{slide_count}</b>\n"
        f"🎨 Dizayn: <b>{template_name}</b>\n\n"
        f"{prepared_by}"
    )



def create_generation_failed_text() -> str:
    return (
        "<b>⚠️ Hozircha taqdimot yaratib bo‘lmadi</b>\n\n"
        "Ayni paytda taqdimot yaratishda vaqtincha cheklov bor. Sarflangan imkoniyat qaytarildi.\n\n"
        "Iltimos, birozdan keyin qayta urinib ko‘ring."
    )



def create_validation_error_text(error_message: str) -> str:
    return f"⚠️ {escape(error_message)}"



def bot_access_blocked_text() -> str:
    return (
        "<b>⛔ Botdan foydalanish hozircha yopiq</b>\n\n"
        "Siz uchun botdan foydalanish vaqtincha cheklangan. Keyinroq qayta urinib ko‘ring yoki aloqa bo‘limiga yozing."
    )



def bot_access_blocked_alert_text() -> str:
    return 'Botdan foydalanish hozircha yopiq. Keyinroq urinib ko‘ring.'


def technical_maintenance_text() -> str:
    return (
        "<b>⚠️ Hozircha xizmatda vaqtincha cheklov bor</b>\n\n"
        "Iltimos, birozdan keyin qayta harakat qilib ko‘ring. "
        "Agar holat davom etsa, aloqa bo‘limi orqali yozishingiz mumkin."
    )


def technical_maintenance_alert_text() -> str:
    return "Hozircha xizmatda vaqtincha cheklov bor. Keyinroq urinib ko‘ring."
