# MathOlimp AI App — FastAPI + HTMX

Migrasi dari aplikasi Streamlit ke FastAPI + Jinja2 + HTMX dengan logika kuis/AI/DB asli dipertahankan sebanyak mungkin. Lapisan kompatibilitas di `main.py` menerjemahkan primitive `st.*` yang dipakai aplikasi menjadi HTML dan request HTMX sehingga kode halaman lama tidak perlu dirombak besar-besaran.

## Struktur

```text
MathOlimp_AI_App/
├── main.py
├── ai_engine.py
├── templates/
│   └── index.html
├── static/
│   └── images/
│       ├── cover.png
│       ├── logo.png
│       └── nexus_logo.png
├── requirements.txt
├── render.yaml
├── .env.example
├── .gitignore
└── README.md
```

## Environment variables

Set minimal:

- `GEMINI_API_KEYS` — satu key atau beberapa key dipisahkan koma/baris baru.
- `DATABASE_URL` — connection string PostgreSQL.
- `SESSION_SECRET` — secret random panjang untuk cookie session.
- `COOKIE_SECURE=true` di production HTTPS; gunakan `false` saat local HTTP.

Jangan commit `.env` atau credential/database password ke GitHub.

## Local run

```bash
python -m pip install -r requirements.txt
uvicorn main:app --reload
```

Buka `http://127.0.0.1:8000`.

## Render.com

`render.yaml` memakai:

```text
Build: pip install -r requirements.txt
Start: uvicorn main:app --host 0.0.0.0 --port $PORT
```

Tambahkan environment variables di Render Dashboard sebelum deployment.

## Catatan migrasi

`ai_engine.py` tidak lagi bergantung pada `st.secrets` atau `st.connection`. API key dibaca dari environment dan PostgreSQL memakai SQLAlchemy, tetapi fungsi AI dan query bisnis yang sudah ada dipertahankan.

Session aplikasi web disimpan server-side berdasarkan session id yang ditandatangani cookie. Ini cukup untuk deployment single-instance. Untuk multi-instance production, session store sebaiknya dipindahkan ke PostgreSQL/Redis.
