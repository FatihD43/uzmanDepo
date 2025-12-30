# UzmanRapor API (FastAPI)

Bu servis, UzmanRapor client uygulamasının SQL Server'a **doğrudan bağlanmadan** çalışması için geliştirilmiştir.

## Kurulum
```bash
cd UZMANRAPOR_API
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

## Çalıştırma
```bash
set UZMANRAPOR_API_TOKEN=degistir_bunu
set UZMANRAPOR_SQL_SERVER=10.30.9.14,1433
set UZMANRAPOR_SQL_DATABASE=UzmanRaporDB
set UZMANRAPOR_SQL_UID=uzmanrapor_login
set UZMANRAPOR_SQL_PWD=xxxx
uvicorn main:app --host 0.0.0.0 --port 8000
```

Alternatif olarak tek parça connection string:
`UZMANRAPOR_SQL_CONN_STR=Driver={SQL Server};Server=...;Database=...;UID=...;PWD=...;`

## Client ayarı
Client'ta env değişkenleri:
- `UZMANRAPOR_API_URL` (ör. `http://sunucu:8000`)
- `UZMANRAPOR_API_TOKEN` (API token)

Not: `/sql` endpoint'i whitelisting + basic token kontrolü içerir. Üretimde ayrıca ağ kısıtları (firewall/VPN) ve daha güçlü kimlik doğrulama önerilir.
