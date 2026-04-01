from __future__ import annotations

from html import escape


_PROGRESS_STAGE_TEXTS = {
    'queued': 'So‘rov navbatga joylandi va ish tartibida kutmoqda.',
    'research': 'Mavzu bo‘yicha mazmun va asosiy faktlar tayyorlanmoqda.',
    'planning': 'Slaydlar tuzilmasi va mazmun bloklari rejalashtirilmoqda.',
    'rendering': 'Taqdimot sahifalari yig‘ilmoqda va formatlanmoqda.',
    'uploading': 'Yakuniy fayl yuborishga tayyorlanmoqda.',
    'done': 'Taqdimot muvaffaqiyatli yakunlandi.',
}


def _progress_bar(percent: int) -> str:
    percent = max(0, min(100, int(percent)))
    filled = max(1, round(percent / 10)) if percent > 0 else 0
    return '█' * filled + '░' * (10 - filled)


def _available_generation_label(user: dict, available_generations: int) -> str:
    if user.get('generation_unlimited'):
        return '♾ <b>Cheksiz</b>'
    if user.get('generation_access_blocked'):
        return '🚫 <b>Vaqtincha cheklangan</b>'
    return f'🎟 <b>{available_generations}</b>'


def main_menu_text(full_name: str) -> str:
    return (
        f"<b>Assalomu alaykum, {escape(full_name)}!</b>\n\n"
        "Ushbu bot mavzu asosida <b>tayyor PPTX taqdimot</b> yaratishga yordam beradi. "
        "Jarayon sodda, tushunarli va bosqichma-bosqich olib boriladi.\n\n"
        "Kerakli bo‘limni tanlash uchun quyidagi tugmalardan birini bosing."
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
        "<b>📊 Shaxsiy holatingiz</b>\n\n"
        f"<u>Asosiy ma’lumotlar</u>\n"
        f"• Ism: <b>{full_name}</b>\n"
        f"• Telegram ID: <code>{telegram_id}</code>\n\n"
        f"<u>Faollik va imkoniyatlar</u>\n"
        f"• Jami urinishlar: <b>{generated_count}</b>\n"
        f"• Muvaffaqiyatli taqdimotlar: <b>{successful}</b>\n"
        f"• Tasdiqlangan takliflar: <b>{referral_count}</b>\n"
        f"• Referral kreditlar: <b>{referral_credits}</b>\n"
        f"• Admin qo‘shgan kreditlar: <b>{bonus_credits}</b>\n"
        f"• Hozirgi foydalanish limiti: {_available_generation_label(user, available_generations)}"
    )



def invite_text(referral_link: str, available_generations: int, referral_count: int, user: dict) -> str:
    return (
        "<b>👥 Taklif qilish bo‘limi</b>\n\n"
        "Do‘stlaringizni ushbu botga taklif qiling. Ular botga kirib, majburiy obunani to‘liq tasdiqlaganidan keyin sizga qo‘shimcha yaratish imkoni beriladi.\n\n"
        f"• Tasdiqlangan takliflar: <b>{referral_count}</b>\n"
        f"• Hozirgi limit: {_available_generation_label(user, available_generations)}\n\n"
        "<u>Shaxsiy taklif havolangiz</u>\n"
        f"<code>{escape(referral_link)}</code>"
    )



def referrals_text(referrals: list[dict]) -> str:
    if not referrals:
        return (
            "<b>👥 Takliflar ro‘yxati</b>\n\n"
            "Hozircha siz orqali ro‘yxatdan o‘tgan foydalanuvchilar aniqlanmadi."
        )

    lines = ["<b>👥 Takliflar ro‘yxati</b>", ""]
    confirmed = 0

    for index, item in enumerate(referrals, start=1):
        status = '✅ tasdiqlangan' if item.get('counted') else '🕓 kutilmoqda'
        if item.get('counted'):
            confirmed += 1
        lines.append(f"{index}. <code>{item.get('invitee_id')}</code> — <b>{status}</b>")

    lines.append('')
    lines.append(f"Jami tasdiqlangan takliflar: <b>{confirmed}</b>")
    return '\n'.join(lines)



def help_text() -> str:
    return (
        "<b>❓ Yordam va foydalanish tartibi</b>\n\n"
        "<u>Bot nima qiladi?</u>\n"
        "• Berilgan mavzu asosida PPTX taqdimot tayyorlaydi\n"
        "• Slaydlar soni va taqdimot tilini tanlash imkonini beradi\n"
        "• Jarayon holatini bosqichma-bosqich ko‘rsatadi\n\n"
        "<u>Qanday foydalaniladi?</u>\n"
        "1. <b>Slayd yaratish</b> bo‘limini tanlang\n"
        "2. Mavzuni yuboring\n"
        "3. Tayyorlagan ismni yuboring\n"
        "4. Slaydlar soni va tilni tanlang\n"
        "5. Ma’lumotlarni tasdiqlang va natijani kuting\n\n"
        "<u>Muhim eslatma</u>\n"
        "• Birinchi yaratish imkoniyati bepul\n"
        "• Keyingi imkoniyatlar referral yoki admin tomonidan berilgan kreditlar orqali ishlaydi\n"
        "• Majburiy obuna faol bo‘lsa, botdan foydalanishdan oldin u to‘liq tasdiqlanishi kerak"
    )



def contact_text(support_contact: str) -> str:
    return (
        "<b>☎️ Aloqa</b>\n\n"
        "Qo‘shimcha savollar, takliflar yoki texnik murojaatlar uchun quyidagi aloqa manzilidan foydalanishingiz mumkin:\n\n"
        f"<b>{escape(support_contact)}</b>"
    )



def subscription_text(channels: list[dict]) -> str:
    if not channels:
        return (
            "<b>📢 Majburiy obuna</b>\n\n"
            "Hozircha faol majburiy kanallar mavjud emas."
        )

    lines = [
        "<b>📢 Botdan foydalanishdan oldin quyidagi kanallarga obuna bo‘ling</b>",
        '',
        "Quyidagi ro‘yxatdagi barcha kanallarga a’zo bo‘lib, so‘ng <b>Obunani tekshirish</b> tugmasini bosing.",
        '',
    ]

    for index, channel in enumerate(channels, start=1):
        title = escape(channel.get('title') or channel.get('username') or f'Kanal {index}')
        lines.append(f"{index}. {title}")

    return '\n'.join(lines)



def subscription_failed_text(channels: list[dict]) -> str:
    lines = [
        "<b>⚠️ Obuna hali to‘liq tasdiqlanmadi</b>",
        '',
        "Quyidagi kanallarga obuna aniqlanmadi:",
        '',
    ]

    for index, channel in enumerate(channels, start=1):
        title = escape(channel.get('title') or channel.get('username') or f'Kanal {index}')
        lines.append(f"{index}. {title}")

    lines.append('')
    lines.append("Barcha kanallarga obuna bo‘lgach, yana <b>Obunani tekshirish</b> tugmasini bosing.")
    return '\n'.join(lines)





def create_generation_blocked_text() -> str:
    return (
        "<b>🚫 Slayd yaratish vaqtincha cheklangan</b>\n\n"
        "Siz uchun taqdimot yaratish imkoniyati administrator tomonidan vaqtincha cheklangan. Batafsil ma’lumot uchun aloqa bo‘limidan foydalaning."
    )

def create_credit_missing_text() -> str:
    return (
        "<b>🎞 Slayd yaratish</b>\n\n"
        "Hozirda sizning hisobingizda yangi taqdimot yaratish uchun yetarli limit aniqlanmadi.\n\n"
        "• Birinchi yaratish — <b>bepul</b>\n"
        "• Keyingi yaratishlar — <b>referral</b> yoki <b>admin qo‘shgan kredit</b> orqali\n\n"
        "Qo‘shimcha limit olish uchun do‘stlaringizni taklif qiling yoki administrator bilan bog‘laning."
    )



def create_topic_prompt_text() -> str:
    return (
        "<b>📝 1-bosqich / 4</b>\n\n"
        "Iltimos, taqdimot uchun <b>mavzuni</b> yuboring.\n\n"
        "Masalan: <i>Sun’iy intellektning ta’limdagi o‘rni</i>"
    )



def create_presenter_prompt_text() -> str:
    return (
        "<b>👤 2-bosqich / 4</b>\n\n"
        "Endi taqdimot muallifi yoki tayyorlagan ismni yuboring."
    )



def create_slide_count_prompt_text() -> str:
    return (
        "<b>📑 3-bosqich / 4</b>\n\n"
        "Taqdimot uchun slaydlar sonini tanlang.\n"
        "Ruxsat etilgan diapazon: <b>6</b> dan <b>15</b> gacha."
    )



def create_language_prompt_text() -> str:
    return (
        "<b>🌐 4-bosqich / 4</b>\n\n"
        "Taqdimot tayyorlanadigan tilni tanlang."
    )



def create_confirmation_text(data: dict) -> str:
    topic = escape(data.get('topic', '-'))
    presenter_name = escape(data.get('presenter_name', '-'))
    slide_count = data.get('slide_count', '-')
    language_name = escape(data.get('language_name', '-'))

    return (
        "<b>✅ Tasdiqlash oynasi</b>\n\n"
        f"• Mavzu: <b>{topic}</b>\n"
        f"• Tayyorlagan ism: <b>{presenter_name}</b>\n"
        f"• Slaydlar soni: <b>{slide_count}</b>\n"
        f"• Til: <b>{language_name}</b>\n\n"
        "Ma’lumotlar to‘g‘ri bo‘lsa, <b>Tasdiqlash</b> tugmasini bosing."
    )



def create_queued_text(data: dict, ahead_count: int) -> str:
    topic = escape(data.get('topic', '-'))
    presenter_name = escape(data.get('presenter_name', '-'))
    slide_count = data.get('slide_count', '-')
    language_name = escape(data.get('language_name', '-'))
    queue_text = (
        'Navbat bo‘sh. So‘rovga tez orada ishlov beriladi.'
        if ahead_count == 0
        else f'Sizdan oldinda <b>{ahead_count}</b> ta so‘rov mavjud.'
    )

    return (
        "<b>📥 So‘rov qabul qilindi</b>\n\n"
        f"• Mavzu: <b>{topic}</b>\n"
        f"• Tayyorlagan ism: <b>{presenter_name}</b>\n"
        f"• Slaydlar soni: <b>{slide_count}</b>\n"
        f"• Til: <b>{language_name}</b>\n\n"
        f"{queue_text}\n"
        "Jarayon bo‘yicha yangilanishlar shu suhbatda ko‘rsatiladi."
    )



def create_already_queued_text(ahead_count: int) -> str:
    queue_text = (
        'Sizning so‘rovingiz hozir faol ishlov jarayonida.'
        if ahead_count == 0
        else f'Sizdan oldinda <b>{ahead_count}</b> ta so‘rov bor.'
    )
    return (
        "<b>⏳ Faol so‘rov mavjud</b>\n\n"
        "Sizda allaqachon navbatga qo‘yilgan yoki ishlanayotgan taqdimot bor.\n\n"
        f"{queue_text}"
    )



def create_generation_progress_text(data: dict, percent: int, stage_key: str) -> str:
    topic = escape(data.get('topic', '-'))
    stage_text = _PROGRESS_STAGE_TEXTS.get(stage_key, 'Jarayon davom etmoqda.')
    bar = _progress_bar(percent)
    return (
        "<b>⚙️ Taqdimot tayyorlanmoqda</b>\n\n"
        f"• Mavzu: <b>{topic}</b>\n"
        f"• Holat: {stage_text}\n\n"
        f"<code>{bar}</code> <b>{percent}%</b>"
    )



def create_generation_success_caption(data: dict) -> str:
    topic = escape(data.get('topic', '-'))
    slide_count = data.get('slide_count', '-')
    return (
        "<b>✅ Taqdimot tayyor</b>\n\n"
        f"• Mavzu: <b>{topic}</b>\n"
        f"• Slaydlar soni: <b>{slide_count}</b>\n\n"
        "Fayl muvaffaqiyatli tayyorlandi."
    )



def create_generation_failed_text() -> str:
    return (
        "<b>⚠️ Taqdimotni yaratishda xatolik yuz berdi</b>\n\n"
        "Sarflangan limit qaytarildi. Iltimos, birozdan keyin qayta urinib ko‘ring."
    )



def create_validation_error_text(error_message: str) -> str:
    return f"⚠️ {escape(error_message)}"



def bot_access_blocked_text() -> str:
    return (
        "<b>⛔ Kirish vaqtincha cheklangan</b>\n\n"
        "Siz uchun botdan foydalanish vaqtincha cheklangan. Qo‘shimcha ma’lumot olish uchun administrator bilan bog‘laning."
    )



def bot_access_blocked_alert_text() -> str:
    return 'Siz uchun botdan foydalanish vaqtincha cheklangan.'


def technical_maintenance_text() -> str:
    return (
        "<b>⚠️ Texnik ishlar</b>\n\n"
        "Texnik ishlar sababli hozirda xizmat ko‘rsatish imkoni yo‘q. "
        "Birozdan keyin harakat qilib ko‘ring.\n\n"
        "Noqulayliklar uchun uzr."
    )


def technical_maintenance_alert_text() -> str:
    return "Texnik ishlar sababli hozircha xizmat ko‘rsatib bo‘lmaydi."
