from __future__ import annotations

from html import escape



def admin_main_menu_text(admin_name: str) -> str:
    return (
        f"<b>🛡 Admin panel</b>\n\n"
        f"Assalomu alaykum, <b>{escape(admin_name)}</b>. Quyidagi bo‘limlar orqali bot statistikasi, foydalanuvchilar, limitlar va ommaviy xabarlarni boshqarishingiz mumkin."
    )



def admin_stats_text(stats: dict) -> str:
    status_counts = stats.get('generation_statuses', {})
    lines = [
        "<b>📈 Umumiy statistika</b>",
        '',
        "<u>Foydalanuvchilar</u>",
        f"• Jami foydalanuvchilar: <b>{stats.get('total_users', 0)}</b>",
        f"• Adminlar: <b>{stats.get('admins', 0)}</b>",
        f"• Obuna bo‘lganlar: <b>{stats.get('subscribed', 0)}</b>",
        f"• Obuna bo‘lmaganlar: <b>{stats.get('unsubscribed', 0)}</b>",
        f"• So‘nggi 24 soat faollar: <b>{stats.get('active_24h', 0)}</b>",
        f"• So‘nggi 7 kun faollar: <b>{stats.get('active_7d', 0)}</b>",
        '',
        "<u>O‘sish</u>",
        f"• Yangi foydalanuvchilar (24 soat): <b>{stats.get('new_24h', 0)}</b>",
        f"• Yangi foydalanuvchilar (7 kun): <b>{stats.get('new_7d', 0)}</b>",
        f"• Yangi foydalanuvchilar (30 kun): <b>{stats.get('new_30d', 0)}</b>",
        '',
        "<u>Referral va limitlar</u>",
        f"• Jami referral natijalari: <b>{stats.get('referral_total', 0)}</b>",
        f"• Admin bergan umumiy kreditlar: <b>{stats.get('manual_bonus_total', 0)}</b>",
        f"• Cheksiz limitli foydalanuvchilar: <b>{stats.get('generation_unlimited', 0)}</b>",
        f"• Generation bloklanganlar: <b>{stats.get('generation_blocked', 0)}</b>",
        f"• Bot kirishi bloklanganlar: <b>{stats.get('bot_blocked', 0)}</b>",
        '',
        "<u>Generation holati</u>",
        f"• Jami urinishlar: <b>{stats.get('total_generated', 0)}</b>",
        f"• Muvaffaqiyatli natijalar: <b>{stats.get('total_successful', 0)}</b>",
        f"• Navbatda: <b>{status_counts.get('queued', 0)}</b>",
        f"• Ishlanmoqda: <b>{status_counts.get('processing', 0)}</b>",
        f"• Tayyorlanganlar: <b>{status_counts.get('done', 0)}</b>",
        f"• Xatolik bilan tugaganlar: <b>{status_counts.get('failed', 0)}</b>",
        f"• Tayyorlanganlar (24 soat): <b>{stats.get('done_24h', 0)}</b>",
        f"• Tayyorlanganlar (7 kun): <b>{stats.get('done_7d', 0)}</b>",
        f"• Xatoliklar (7 kun): <b>{stats.get('failed_7d', 0)}</b>",
    ]

    recent_failures = stats.get('recent_failures') or []
    if recent_failures:
        lines.extend(['', '<u>Oxirgi xatoliklar</u>'])
        for item in recent_failures[:5]:
            lines.append(
                f"• <code>{item.get('telegram_id')}</code> — {escape(str(item.get('error') or 'noma’lum xatolik'))[:80]}"
            )
    return '\n'.join(lines)



def admin_rating_text(ratings: dict) -> str:
    top_referrers = ratings.get('top_referrers') or []
    top_generators = ratings.get('top_generators') or []
    lines = ["<b>🏆 Reyting va top foydalanuvchilar</b>", '']

    lines.append('<u>Referral bo‘yicha top 10</u>')
    if top_referrers:
        for index, user in enumerate(top_referrers, start=1):
            name = escape(user.get('full_name') or str(user.get('telegram_id')))
            lines.append(f"{index}. <b>{name}</b> — {user.get('referral_count', 0)} ta referral")
    else:
        lines.append('• Ma’lumot topilmadi')

    lines.extend(['', '<u>Generation bo‘yicha top 10</u>'])
    if top_generators:
        for index, user in enumerate(top_generators, start=1):
            name = escape(user.get('full_name') or str(user.get('telegram_id')))
            lines.append(
                f"{index}. <b>{name}</b> — {user.get('successful_generations', 0)} ta muvaffaqiyatli / {user.get('generated_count', 0)} ta urinish"
            )
    else:
        lines.append('• Ma’lumot topilmadi')

    return '\n'.join(lines)



def admin_user_search_prompt_text() -> str:
    return (
        "<b>🔎 Foydalanuvchi qidiruvi</b>\n\n"
        "Telegram ID, username yoki ism bo‘yicha qidiruv yuboring.\n\n"
        "Misollar:\n"
        "• <code>123456789</code>\n"
        "• <code>@username</code>\n"
        "• <code>Ali Valiyev</code>"
    )



def admin_search_results_text(query: str, results: list[dict]) -> str:
    if not results:
        return (
            "<b>🔎 Qidiruv natijasi</b>\n\n"
            f"<b>{escape(query)}</b> bo‘yicha foydalanuvchi topilmadi."
        )

    lines = [f"<b>🔎 Qidiruv natijasi</b>", '', f"So‘rov: <b>{escape(query)}</b>", '']
    for index, user in enumerate(results, start=1):
        name = escape(user.get('full_name') or 'Noma’lum')
        username = user.get('username') or '—'
        username_text = f"@{escape(username).lstrip('@')}" if username != '—' else '—'
        lines.append(f"{index}. <b>{name}</b> — <code>{user.get('telegram_id')}</code> — {username_text}")
    lines.append('')
    lines.append('Pastdagi tugmalardan foydalanuvchini tanlang.')
    return '\n'.join(lines)



def admin_user_card_text(card: dict) -> str:
    user = card['user']
    available = card['available_generations']
    active_job = card.get('active_job')
    queue_ahead = int(card.get('queue_ahead', 0) or 0)

    available_label = '♾ Cheksiz' if user.get('generation_unlimited') else str(available)
    username = user.get('username') or '—'
    status_flags = []
    status_flags.append('admin' if user.get('is_admin') else 'user')
    status_flags.append('subscribed' if user.get('subscription_verified') else 'unsubscribed')
    if user.get('generation_access_blocked'):
        status_flags.append('generation-blocked')
    if user.get('bot_access_blocked'):
        status_flags.append('bot-blocked')

    lines = [
        "<b>👤 Foydalanuvchi kartasi</b>",
        '',
        f"• Ism: <b>{escape(user.get('full_name') or 'Noma’lum')}</b>",
        f"• Telegram ID: <code>{user.get('telegram_id')}</code>",
        f"• Username: <b>{escape(username)}</b>",
        f"• Holat: <b>{', '.join(status_flags)}</b>",
        '',
        "<u>Limitlar va referral</u>",
        f"• Hozirgi foydalanish limiti: <b>{available_label}</b>",
        f"• Referral kreditlar: <b>{user.get('referral_credits', 0)}</b>",
        f"• Admin qo‘shgan kreditlar: <b>{user.get('bonus_generation_credits', 0)}</b>",
        f"• Referral soni: <b>{user.get('referral_count', 0)}</b>",
        '',
        "<u>Generation statistikasi</u>",
        f"• Jami urinishlar: <b>{user.get('generated_count', 0)}</b>",
        f"• Muvaffaqiyatli natijalar: <b>{user.get('successful_generations', 0)}</b>",
        '',
        "<u>Vaqt ma’lumotlari</u>",
        f"• Ro‘yxatdan o‘tgan: <b>{escape(str(user.get('created_at') or '—'))}</b>",
        f"• Oxirgi faollik: <b>{escape(str(user.get('last_active_at') or '—'))}</b>",
    ]

    if active_job:
        status = escape(active_job.get('status', 'unknown'))
        lines.extend(
            [
                '',
                "<u>Faol so‘rov</u>",
                f"• Status: <b>{status}</b>",
                f"• Navbat oldi: <b>{queue_ahead}</b>",
            ]
        )

    return '\n'.join(lines)



def admin_credit_prompt_text(user: dict, operation: str) -> str:
    operation_title = 'qo‘shish' if operation == 'add' else 'ayirish'
    return (
        f"<b>🎟 Kredit {operation_title}</b>\n\n"
        f"Foydalanuvchi: <b>{escape(user.get('full_name') or str(user.get('telegram_id')))}</b>\n"
        f"Joriy admin kreditlari: <b>{user.get('bonus_generation_credits', 0)}</b>\n\n"
        "Iltimos, butun son yuboring. Masalan: <code>3</code>"
    )



def admin_broadcast_menu_text() -> str:
    return (
        "<b>📣 Ommaviy xabarlar</b>\n\n"
        "Quyidagi auditoriyalardan birini tanlang. Keyin xabar matni yoki media yuborib, kerak bo‘lsa inline tugmalar qo‘shishingiz mumkin."
    )



def admin_broadcast_content_prompt_text(audience_label: str, audience_count: int) -> str:
    return (
        "<b>✍️ Xabar kontenti</b>\n\n"
        f"Tanlangan auditoriya: <b>{escape(audience_label)}</b>\n"
        f"Qamrov: <b>{audience_count}</b> ta foydalanuvchi\n\n"
        "Endi yuboriladigan xabarni jo‘nating.\n"
        "Qo‘llab-quvvatlanadi: <b>matn, rasm, video, GIF, document</b>.\n"
        "Matn ichida <b>HTML teglari</b> ishlatishingiz mumkin: <code>&lt;b&gt;</code>, <code>&lt;i&gt;</code>, <code>&lt;u&gt;</code>, <code>&lt;code&gt;</code>."
    )



def admin_broadcast_buttons_prompt_text() -> str:
    return (
        "<b>🔘 Inline tugmalar</b>\n\n"
        "Har bir tugmani yangi qatordan quyidagi formatda yuboring:\n"
        "<code>Tugma matni | https://example.com</code>\n"
        "yoki\n"
        "<code>Tugma matni | callback:help</code>\n\n"
        "Ruxsat etilgan callback qiymatlar: <code>main</code>, <code>status</code>, <code>invite</code>, <code>help</code>, <code>contact</code>\n\n"
        "Agar tugma kerak bo‘lmasa, <b>O‘tkazib yuborish</b> tugmasini bosing."
    )



def admin_broadcast_preview_text(audience_label: str, audience_count: int, buttons_count: int, draft: dict) -> str:
    kind = escape(draft.get('kind', 'text'))
    return (
        "<b>🧪 Xabar preview tayyor</b>\n\n"
        f"• Auditoriya: <b>{escape(audience_label)}</b>\n"
        f"• Qamrov: <b>{audience_count}</b> ta foydalanuvchi\n"
        f"• Kontent turi: <b>{kind}</b>\n"
        f"• Tugmalar soni: <b>{buttons_count}</b>\n\n"
        "Pastdagi boshqaruv tugmalari orqali matnni almashtirish, tugmalarni yangilash, test yuborish yoki final jo‘natishni bajarishingiz mumkin."
    )



def admin_broadcast_result_text(audience_label: str, result: dict) -> str:
    return (
        "<b>✅ Ommaviy xabar yakunlandi</b>\n\n"
        f"• Auditoriya: <b>{escape(audience_label)}</b>\n"
        f"• Qamrab olingan: <b>{result.get('processed', 0)}</b>\n"
        f"• Muvaffaqiyatli yuborilgan: <b>{result.get('success', 0)}</b>\n"
        f"• Xatolik bilan yakunlangan: <b>{result.get('failed', 0)}</b>"
    )



def admin_export_text() -> str:
    return (
        "<b>📤 Eksport bo‘limi</b>\n\n"
        "Kerakli auditoriyani tanlang va CSV yoki XLSX formatda foydalanuvchilar ro‘yxatini yuklab oling."
    )



def admin_export_ready_text(audience_label: str, fmt: str, count: int) -> str:
    return (
        "<b>📎 Eksport fayli tayyor</b>\n\n"
        f"• Auditoriya: <b>{escape(audience_label)}</b>\n"
        f"• Format: <b>{escape(fmt.upper())}</b>\n"
        f"• Yozilgan foydalanuvchilar soni: <b>{count}</b>"
    )



def admin_simple_result_text(message: str) -> str:
    return f"<b>✅ Amal bajarildi</b>\n\n{escape(message)}"
