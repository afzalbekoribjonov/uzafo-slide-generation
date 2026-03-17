# Slide Bot Foundation

Bu loyiha `aiogram 3` va `MongoDB` asosida qurilgan Telegram bot skeleti.

Hozirgi bosqichda quyidagilar tayyor:
- `/start` va deeplink (`/start <inviter_id>`) ishlaydi
- foydalanuvchi MongoDB ga saqlanadi
- asosiy menyu inline keyboard bilan ishlaydi
- `Holat, Taklif qilish` bo'limi ishlaydi
- `Men taklif qilganlar` bo'limi ishlaydi
- `Yordam, Aloqa` bo'limi ishlaydi
- referral yozuvi (start bosqichida) tayyorlangan
- keyingi bosqichlar uchun `service/repository` arxitekturasi tayyor

## Pipenv bilan o'rnatish

### 1) Pipenv o'rnatish
Agar sizda `pipenv` o'rnatilmagan bo'lsa:

```bash
pip install pipenv
```

### 2) Virtual muhit va kutubxonalarni o'rnatish
Loyiha papkasida:

```bash
pipenv install
```

### 3) `.env` fayl tayyorlash

```bash
cp .env.example .env
```

Windows uchun:

```powershell
copy .env.example .env
```

So'ng `.env` ichida token va MongoDB ma'lumotlarini to'ldiring.

## Ishga tushirish

### Variant 1
```bash
pipenv run python -m app.main
```

### Variant 2
`Pipfile` ichidagi script orqali:

```bash
pipenv run start
```

### Variant 3
Shell ichiga kirib ishga tushirish

```bash
pipenv shell
python -m app.main
```

## Muhim fayllar
- `app/main.py` — bot entrypoint
- `app/handlers/user/start.py` — `/start` va deeplink logikasi
- `app/handlers/user/menu.py` — inline menyular
- `app/services/` — biznes logika
- `app/repositories/` — MongoDB bilan ishlash
- `app/keyboards/user.py` — inline tugmalar

## Keyingi bosqichlar
1. Majburiy obuna (bir nechta kanal)
2. Admin panel va kanal boshqaruvi
3. Slayd yaratish oqimi
4. PPTX generation service


## Render uchun webhook deploy

Bu loyiha endi ikki rejimni qo‘llaydi:
- `APP_MODE=polling` — lokal test uchun
- `APP_MODE=webhook` — Render Web Service uchun

### Tavsiya etilgan Render sozlamalari
- **Build Command:** `pip install -r requirements.txt`
- **Start Command:** `python -m app.main`
- **Health Check Path:** `/healthz`

### Render uchun kerakli env qiymatlar
`.env` yoki Render Environment bo‘limiga quyidagilarni kiriting:

```env
APP_MODE=webhook
WEBHOOK_BASE_URL=https://your-service-name.onrender.com
WEBHOOK_PATH=/telegram/webhook
WEBHOOK_SECRET=change_me_please
```

`WEBHOOK_BASE_URL` qiymati Render servisining tashqi HTTPS manzili bo‘lishi kerak.

### Muhim eslatmalar
- Webhook rejimida bot `PORT` yoki `WEB_SERVER_PORT` ga quloq soladi.
- Bot ishga tushganda webhook avtomatik o‘rnatiladi.
- `/healthz` endpoint Render health check uchun tayyor.
- PPTX generation logikasi o‘zgartirilmagan.
