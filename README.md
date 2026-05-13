# Slide Generator Bot

Ushbu loyiha sun'iy intellekt (Gemini AI) yordamida professional PowerPoint (`.pptx`) va PDF taqdimotlar yaratish uchun mo'ljallangan Telegram bot va WebApp tizimidir.

## 🚀 Asosiy Imkoniyatlar
- **Tezkor yaratish:** Mavzu, muallif va slaydlar sonini kiritib, 8 xil professional dizayndagi taqdimotni generatsiya qilish.
- **Sun'iy Intellekt:** Gemini AI orqali mavzu bo'yicha strukturaviy reja va matnlarni avtomatik shakllantirish.
- **WebApp Integratsiyasi:** Telegram Mini App (TMA) orqali qulay va vizual interfeysda boshqarish.
- **Ko'p formatli yuklash:** Tayyor fayllarni `.pptx` va `.pdf` formatlarida olish imkoniyati.
- **Referral Tizimi:** Do'stlarni taklif qilish orqali yangi generatsiya imkoniyatlarini (kredit) qo'lga kiritish.
- **Real-time Status:** Slaydlar yaratilish jarayonini real vaqt rejimida kuzatib borish.

## 🛠 Texnologik stek
### Backend
- **Python 3.11**
- **aiogram 3** (Telegram bot API)
- **aiohttp** (WebApp va API uchun)
- **MongoDB** (Foydalanuvchilar va generatsiyalar tarixi)
- **python-pptx** (Slayd yaratish uchun)
- **Google GenAI SDK** (AI integratsiyasi)

### Frontend (WebApp)
- **React + TypeScript**
- **Vite** (Build tool)
- **Tailwind CSS** (Dizayn tizimi)
- **Framer Motion** (Interaktiv animatsiyalar)

### Infrastruktura
- **Docker** (Multi-stage build: Backend va Frontend yagona konteynerda)
- **Render.com** (Deployment)

## 🏗 Arxitektura
Loyihada `Service/Repository` patterndan foydalanilgan. Backend va Frontend yagona serverda birlashtirilgan bo'lib, Python backend ham API so'rovlarini, ham statik WebApp fayllarini (Single Page Application) tarqatadi.

## 👨‍💻 Dasturchi
**Afzalbek Oribjonov** ([uzafo.uz](https://uzafo.uz))

---
*Loyiha taqdimotlar tayyorlash jarayonini avtomatlashtirish va vaqtni tejash uchun ishlab chiqilgan.*
