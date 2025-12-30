# app/db_name.py
from __future__ import annotations

from app.site_config import get_db_name

# Diğer tüm modüller aynı şekilde "from app.db_name import DB_NAME" kullanmaya devam edebilir.
DB_NAME: str = get_db_name()
