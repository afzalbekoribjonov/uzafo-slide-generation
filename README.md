# Slide Generator Bot

Ushbu loyiha foydalanuvchilarga qisqa vaqt ichida professional ko'rinishdagi PowerPoint (`.pptx`) va PDF taqdimotlarni avtomatik yaratish imkonini beruvchi yagona ekotizimdir. Telegram bot qulayligi va WebApp interfeysining estetikasi bir nuqtada birlashtirilgan.

## 🌟 Loyihaning asosiy afzalliklari
- **Yagona ekotizim:** Telegram bot va WebApp bir serverda, yagona bazada ishlaydi. Bu foydalanuvchi tajribasini uzluksiz qiladi.
- **Dizayn erkinligi:** 8 xil professional shablonlar to'plami. Har bir shablon ranglar palitrasi, shriftlar va slayd joylashuvi bo'yicha noyob.
- **Tezkor va qulay:** Mavzu va muallif ma'lumotlarini kiritish orqali bir necha soniyada tayyor taqdimot.
- **Avtomatlashtirilgan PDF konvertatsiyasi:** Tayyor taqdimotni darhol PDF formatida ham olish imkoniyati.
- **Referral tizimi:** Do'stlarni taklif qilish orqali bepul imkoniyatlarni (kredit) oshirish imkoniyati.
- **Real-time kuzatuv:** Slaydlar yaratilish jarayonini WebApp orqali real vaqt rejimida kuzatish mumkin.

## 🛠 Texnik yechimlar
- **Unified Deployment:** Backend va Frontend yagona Docker konteynerida. Statik fayllar (WebApp) Python backend tomonidan taqdim etiladi.
- **Micro-service arxitektura:** `Service/Repository` patterndan foydalangan holda kod bazasi modulli va oson kengaytiriladigan qilingan.
- **Telegram Mini App (TMA):** Bot ichida to'liq interaktiv interfeys.

## 🚀 Qanday ishga tushirish kerak?

### 1. Talablar
- Docker va Docker Compose
- MongoDB bazasi
- Telegram Bot Token

### 2. Sozlash
Fayllar tarkibidagi `.env.example` faylini `.env` deb nusxalash va zaruriy o'zgaruvchilarni kiritish:
```bash
cp .env.example .env
# .env faylini ochib, kerakli ma'lumotlarni kiriting
```

### 3. Ishga tushirish (Docker)
Loyihani lokal yoki serverda Docker yordamida quyidagi buyruq bilan ishga tushirish mumkin:
```bash
docker build -t slide-generator .
docker run -p 10000:10000 --env-file .env slide-generator
```

## 👨‍💻 Dasturchi
**Afzalbek Oribjonov** ([uzafo.uz](https://uzafo.uz))

---
*Loyiha taqdimotlar tayyorlash jarayonini avtomatlashtirish va vaqtni tejash uchun ishlab chiqilgan.*
