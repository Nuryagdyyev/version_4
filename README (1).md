# Akademik Işler Boty 🎓

Telegram boty — Referat we Doklad taýarlaýar (Word .docx görnüşinde, rus dilinde).

## Faýllar

```
├── bot.py           # Esasy bot kody
├── template.docx    # Kapylyk sahypasy şablony
├── requirements.txt # Python garaşlylyklary
├── render.yaml      # Render.com awtomatik sazlama
└── .gitignore
```

## Deploy (Render.com)

1. Bu repo-ny GitHub-a ýükläň
2. [render.com](https://render.com) → New → Web Service → GitHub repo saýlaň
3. Environment Variables goşuň (aşakda görkezilen)
4. Disks → `/data` Mount Path, 1 GB
5. Deploy

## Environment Variables

| Açar | Düşündiriş | Mysal |
|------|-----------|-------|
| `BOT_TOKEN` | Telegram bot tokeni (@BotFather-dan) | `123456:ABC...` |
| `DEEPSEEK_API_KEY` | DeepSeek API açary | `sk-...` |
| `DB_PATH` | SQLite DB ýoly (Disk bilen) | `/data/bot.db` |
| `REDIS_URL` | (Islege bagly) Redis URL | `redis://...` |

## Admin komandalar

- `/admin` — kömek
- `/stats` — statistika
- `/users` — /start basan ulanyjylar sanawy
- `/orders` — garaşyp duran sargytlar
- `/send <user_id>` — sargyt faýlyny iber

## Bahalar

- 📄 Referat — 299 ₽
- 🎤 Doklad — 299 ₽
- 🎁 1-nji sargyt MUGT
