from __future__ import annotations
from datetime import datetime
from zoneinfo import ZoneInfo
from typing import List, Dict

import ast
import base64
import re
import secrets
import hashlib
import io
import pickle

import pandas as pd
import pyodbc
import zlib


# ============================================================
#  SQL SERVER BAĞLANTISI
# ============================================================

SQL_CONN_STR = (
    "Driver={SQL Server};"
    "Server=10.30.9.14,1433;"
    "Database=UzmanRaporDB;"
    "UID=uzmanrapor_login;"
    "PWD=03114080Ww.;"
)


def _sql_conn():
    """UzmanRaporDB bağlantısı."""
    return pyodbc.connect(SQL_CONN_STR)


#  APP META TABLOSU (GENEL ANAHTAR/DEĞER)
#    - last_update       : planlama güncellik zamanı (ISO datetime)
#    - last_username     : en son giriş yapan kullanıcı adı

def _ensure_meta_table() -> None:
    """
    AppMeta tablosunun varlığını garanti eder.

    AppMeta:
      MetaKey   NVARCHAR(50) PRIMARY KEY
      MetaValue NVARCHAR(MAX) NULL
      UpdatedAt DATETIME2(0) NOT NULL DEFAULT SYSUTCDATETIME()
    """
    sql = """
    IF OBJECT_ID('dbo.AppMeta', 'U') IS NULL
    BEGIN
        CREATE TABLE dbo.AppMeta (
            MetaKey   nvarchar(50) NOT NULL PRIMARY KEY,
            MetaValue nvarchar(max) NULL,
            UpdatedAt datetime2(0) NOT NULL
                CONSTRAINT DF_AppMeta_UpdatedAt DEFAULT (SYSUTCDATETIME())
        );
    END
    """
    with _sql_conn() as c:
        cur = c.cursor()
        cur.execute(sql)
        c.commit()


def _meta_get(key: str) -> str | None:
    """AppMeta içinden tek bir anahtarın değerini döner."""
    _ensure_meta_table()
    with _sql_conn() as c:
        cur = c.cursor()
        cur.execute("SELECT MetaValue FROM dbo.AppMeta WHERE MetaKey = ?;", (key,))
        row = cur.fetchone()
    if not row:
        return None
    return row[0]


def _meta_set(key: str, value: str | None) -> None:
    """AppMeta içine tek anahtar/değer yazar (upsert)."""
    _ensure_meta_table()
    with _sql_conn() as c:
        cur = c.cursor()
        sql = """
        MERGE dbo.AppMeta AS target
        USING (SELECT ? AS MetaKey) AS src
            ON target.MetaKey = src.MetaKey
        WHEN MATCHED THEN
            UPDATE SET MetaValue = ?, UpdatedAt = SYSUTCDATETIME()
        WHEN NOT MATCHED THEN
            INSERT (MetaKey, MetaValue, UpdatedAt)
            VALUES (src.MetaKey, ?, SYSUTCDATETIME());
        """
        cur.execute(sql, (key, value, value))
        c.commit()


#  NOT KURALLARI (Tamamen SQL: dbo.NoteRules veya AppMeta üzerinden saklama)


def _note_rules_table_exists() -> bool:
    """NoteRules tablosu erişilebilir mi?"""
    sql = """
    SELECT 1
    FROM sys.objects
    WHERE object_id = OBJECT_ID('dbo.NoteRules')
      AND type = 'U';
    """
    try:
        with _sql_conn() as c:
            cur = c.cursor()
            cur.execute(sql)
            return cur.fetchone() is not None
    except Exception:
        return False

def _decode_rules_from_meta(raw: str | bytes | None) -> list[dict]:
    if not raw:
        return []
    try:
        blob = base64.b64decode(raw)
        obj = pickle.loads(blob)
        if isinstance(obj, list):
            return [r for r in obj if isinstance(r, dict)]
    except Exception:
        pass
    return []


def load_rules() -> list[dict]:
    """
    Not kurallarını SQL'den okur.

    Öncelik: AppMeta.note_rules (base64 + pickle ile saklanmış liste)
    Geriye dönük: NoteRules tablosu varsa buradan okumayı dener.
    """
    meta_rules = _decode_rules_from_meta(_meta_get("note_rules"))
    if meta_rules:
        return meta_rules

    if not _note_rules_table_exists():
        return []

    rules: list[dict] = []
    try:
        with _sql_conn() as c:
            cur = c.cursor()
            cur.execute("SELECT RuleData FROM dbo.NoteRules ORDER BY Id;")
            rows = cur.fetchall()

        for row in rows:
            blob = row[0]
            if blob is None:
                continue
            try:
                rule = pickle.loads(blob)
                if isinstance(rule, dict):
                    rules.append(rule)
            except Exception:
                continue
    except Exception:
        pass
    return rules


def save_rules(rules: list[dict]) -> None:
    """
    Not kurallarını AppMeta.note_rules anahtarına yazar.
    Ayrıca NoteRules tablosu varsa (ve erişilebiliyorsa) orayı da günceller.
    """
    if not isinstance(rules, list):
        return

    cleaned = [r for r in rules if isinstance(r, dict)]

    # --- AppMeta'ya base64 pickle olarak yaz ---
    try:
        if cleaned:
            blob = pickle.dumps(cleaned)
            payload = base64.b64encode(blob).decode("ascii")
            _meta_set("note_rules", payload)
        else:
            _meta_set("note_rules", None)
    except Exception:
        pass

    # --- Eğer NoteRules tablosu erişilebilir ise aynı veriyi oraya da yaz ---

    if not _note_rules_table_exists():
        return

    try:
        with _sql_conn() as c:
            cur = c.cursor()
            cur.execute("DELETE FROM dbo.NoteRules;")

            for rule in cleaned:
                try:
                    blob = pickle.dumps(rule)
                except Exception:
                    continue
                cur.execute(
                    "INSERT INTO dbo.NoteRules (RuleData) VALUES (?);",
                    (pyodbc.Binary(blob),),
                )
            c.commit()
    except Exception:
        pass
# ============================================================
#  SON GÜNCELLEME (Planlama tıklanınca kaydedilen zaman)
#    - AppMeta.MetaKey = 'last_update'
# ============================================================

def load_last_update() -> datetime | None:
    raw = _meta_get("last_update")
    if not raw:
        return None
    try:
        dt = datetime.fromisoformat(raw)
        # naive ise İstanbul TZ ata
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=ZoneInfo("Europe/Istanbul"))
        return dt
    except Exception:
        return None


def save_last_update(dt: datetime) -> None:
    # dt naive ise İstanbul TZ'li kabul et
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=ZoneInfo("Europe/Istanbul"))
    iso = dt.isoformat()
    _meta_set("last_update", iso)


# ============================================================
#  SNAPSHOT (Dinamik & Running) – SQL TABLOSU
# ============================================================

# ============================================================
#  SNAPSHOT (Dinamik & Running) – SQL TABLOSU (HEX STRING)
# ============================================================

# ============================================================
#  SNAPSHOT (Dinamik & Running) – SQL TABLOSU (HEX + ZLIB)
# ============================================================

def _ensure_snapshot_table() -> None:
    """
    Snapshots tablosu:
      Name   : 'dinamik' veya 'running' vb
      DataHex: pandas pickle'ın ZLIB sıkıştırılmış hex string hâli (nvarchar(max))
    """
    sql = """
    IF OBJECT_ID('dbo.Snapshots', 'U') IS NULL
    BEGIN
        CREATE TABLE dbo.Snapshots (
            Name      nvarchar(50) NOT NULL PRIMARY KEY,
            DataHex   nvarchar(max) NULL,
            UpdatedAt datetime2 NOT NULL CONSTRAINT DF_Snapshots_UpdatedAt DEFAULT (sysdatetime())
        );
    END
    """
    try:
        with _sql_conn() as c:
            cur = c.cursor()
            cur.execute(sql)
            c.commit()
    except Exception:
        # tablo yoksa sonraki save çağrısında tekrar denenecek
        pass


def save_df_snapshot(df: pd.DataFrame | None, which: str) -> None:
    """
    Dinamik / Running DataFrame'lerini SQL'e pickle + ZLIB + HEX olarak yazar.
    which: 'dinamik' veya 'running' vb.
    """
    if df is None:
        return

    _ensure_snapshot_table()

    try:
        # 1) DataFrame -> pickle bytes
        buf = io.BytesIO()
        df.to_pickle(buf)
        raw_bytes = buf.getvalue()

        # 2) ZLIB ile sıkıştır
        compressed = zlib.compress(raw_bytes, level=9)

        # 3) HEX string'e çevir
        hex_str = compressed.hex()

        print(
            f"[SNAPSHOT] {which}: {len(df)} satır kaydedildi, "
            f"raw={len(raw_bytes)} byte, compressed={len(compressed)} byte"
        )

        with _sql_conn() as c:
            cur = c.cursor()
            # Önce var olan kaydı sil
            cur.execute("DELETE FROM dbo.Snapshots WHERE Name = ?;", (which,))
            # Sonra yeni kaydı ekle (hex string olarak)
            cur.execute(
                "INSERT INTO dbo.Snapshots (Name, DataHex) VALUES (?, ?);",
                (which, hex_str),
            )
            c.commit()
    except Exception as e:
        print(f"[SNAPSHOT] {which}: KAYIT HATASI -> {e!r}")


def load_df_snapshot(which: str) -> pd.DataFrame | None:
    """
    Dinamik / Running snapshot'ı SQL'den okur.
    Kayıt yoksa None döner.
    """
    _ensure_snapshot_table()

    try:
        with _sql_conn() as c:
            cur = c.cursor()
            cur.execute("SELECT DataHex FROM dbo.Snapshots WHERE Name = ?;", (which,))
            row = cur.fetchone()

        if not row or row[0] is None:
            return None

        hex_str = row[0]

        # 1) HEX -> sıkıştırılmış bytes
        compressed = bytes.fromhex(hex_str)
        # 2) ZLIB decompress -> raw pickle bytes
        raw_bytes = zlib.decompress(compressed)

        buf = io.BytesIO(raw_bytes)
        df = pd.read_pickle(buf)
        if isinstance(df, pd.DataFrame):
            return df
    except Exception as e:
        print(f"[SNAPSHOT] {which}: YÜKLEME HATASI -> {e!r}")

    return None

# ============================================================
#  KULLANICI VARSAYILANI (login ekranındaki son kullanıcı)
#    - AppMeta.MetaKey = 'last_username'
# ============================================================

def get_username_default() -> str:
    """
    Login ekranında görünen varsayılan kullanıcı adı.
    Artık SQL'den (AppMeta.last_username) okunuyor.
    """
    val = _meta_get("last_username")
    if val:
        return str(val)
    return "Anonim"


def set_username_default(name: str) -> None:
    """
    Giriş yapıldığında en son kullanılan kullanıcı adını yazar.
    """
    name = (name or "").strip()
    if not name:
        return
    _meta_set("last_username", name)


# ============================================================
#  LOGIN SİSTEMİ – AppUsers (SQL)
# ============================================================

def hash_password(password: str, salt: str) -> str:
    """Parola + salt için SHA-256 hash."""
    payload = f"{salt}:{password}".encode("utf-8")
    return hashlib.sha256(payload).hexdigest()


def _ensure_app_users_table() -> None:
    sql = """
    IF OBJECT_ID('dbo.AppUsers', 'U') IS NULL
    BEGIN
        CREATE TABLE dbo.AppUsers (
            Username     nvarchar(50) NOT NULL PRIMARY KEY,
            Salt         nvarchar(64) NOT NULL,
            PasswordHash nvarchar(64) NOT NULL,
            Permissions  nvarchar(max) NULL,      -- Virgül ile ayrılmış izinler (örn: admin,read,write)
            IsActive     bit NOT NULL CONSTRAINT DF_AppUsers_IsActive DEFAULT (1),
            CreatedAt    datetime2(0) NOT NULL CONSTRAINT DF_AppUsers_CreatedAt DEFAULT (SYSUTCDATETIME())
        );
    END
    """
    with _sql_conn() as c:
        cur = c.cursor()
        cur.execute(sql)
        c.commit()


def ensure_user_db() -> None:
    """
    Kullanıcı tablosunun varlığını ve en az 1 admin kullanıcısını garanti eder.
    Varsayılan: admin / admin
    """
    try:
        _ensure_app_users_table()

        with _sql_conn() as c:
            cur = c.cursor()
            cur.execute("SELECT COUNT(*) FROM dbo.AppUsers;")
            count = cur.fetchone()[0] or 0

            if count == 0:
                # Varsayılan admin␊
                salt = secrets.token_hex(16)
                pwd_hash = hash_password("admin", salt)
                perms = "admin,read,write"

                cur.execute(
                    """
                    INSERT INTO dbo.AppUsers (Username, Salt, PasswordHash, Permissions, IsActive)
                    VALUES (?, ?, ?, ?, 1);
                    """,
                    ("admin", salt, pwd_hash, perms),
                )
            c.commit()
    except Exception:
        # En kötü ihtimalle login ekranı boş kullanıcı listesi ile açılır
        pass


def load_users() -> list[dict]:
    """
    Uygulama kullanıcılarını SQL'den okur.
    Geri dönüş örneği:
      [
        {
          "username": "admin",
          "salt": "...",
          "password_hash": "...",
          "permissions": ["admin","read","write"],
          "is_active": True,
          "created_at": "2025-11-14T..."
        },
        ...
      ]
    """
    ensure_user_db()
    users: list[dict] = []

    try:
        with _sql_conn() as c:
            cur = c.cursor()
            cur.execute(
                "SELECT Username, Salt, PasswordHash, Permissions, IsActive, CreatedAt "
                "FROM dbo.AppUsers;"
            )
            rows = cur.fetchall()

        for row in rows:
            username = row[0]
            salt = row[1]
            pwd_hash = row[2]
            perms_raw = row[3]
            is_active = bool(row[4])
            created_at = row[5]

            if isinstance(perms_raw, str):
                parsed: list[str] | str
                try:
                    tmp = ast.literal_eval(perms_raw)
                    parsed = tmp if isinstance(tmp, list) else perms_raw
                except Exception:
                    parsed = perms_raw
                if isinstance(parsed, list):
                    perms = [str(p).strip() for p in parsed if str(p).strip()]
                else:
                    perms = [p.strip() for p in parsed.split(",") if p.strip()]
            else:
                perms = []

            users.append(
                {
                    "username": username,
                    "salt": salt,
                    "password_hash": pwd_hash,
                    "permissions": perms,
                    "is_active": is_active,
                    "created_at": created_at.isoformat()
                    if isinstance(created_at, datetime)
                    else str(created_at),
                }
            )
    except Exception:
        pass

    return users


def save_users(users: list[dict]) -> None:
    """
    Kullanıcı listesini tamamen SQL'e yazar.
    Önce tüm satırlar silinir, sonra verilen liste baştan eklenir.
    """
    ensure_user_db()
    if not isinstance(users, list):
        return

    try:
        with _sql_conn() as c:
            cur = c.cursor()
            # Tüm kullanıcıları temizle
            cur.execute("DELETE FROM dbo.AppUsers;")

            for u in users:
                username = str(u.get("username", "")).strip()
                salt = str(u.get("salt", "")).strip()
                pwd_hash = str(u.get("password_hash", "")).strip()
                perms = u.get("permissions", [])

                if not username or not salt or not pwd_hash:
                    continue

                if isinstance(perms, list):
                    perms_raw = ",".join([str(p).strip() for p in perms if str(p).strip()])
                else:
                    perms_raw = str(perms)

                is_active = u.get("is_active", True)
                is_active_bit = 1 if is_active else 0

                cur.execute(
                    """
                    INSERT INTO dbo.AppUsers (Username, Salt, PasswordHash, Permissions, IsActive)
                    VALUES (?, ?, ?, ?, ?);
                    """,
                    (username, salt, pwd_hash, perms_raw, is_active_bit),
                )

            c.commit()
    except Exception:
        pass


def find_user(username: str) -> dict | None:
    """Tek bir kullanıcıyı (varsa) döner."""
    ensure_user_db()
    try:
        with _sql_conn() as c:
            cur = c.cursor()
            cur.execute(
                "SELECT Username, Salt, PasswordHash, Permissions, IsActive, CreatedAt "
                "FROM dbo.AppUsers WHERE Username = ?;",
                (username,),
            )
            row = cur.fetchone()
        if not row:
            return None

        perms_raw = row[3]
        perms_raw = row[3]
        if isinstance(perms_raw, str):
            try:
                tmp = ast.literal_eval(perms_raw)
                parsed = tmp if isinstance(tmp, list) else perms_raw
            except Exception:
                parsed = perms_raw
            if isinstance(parsed, list):
                perms = [str(p).strip() for p in parsed if str(p).strip()]
            else:
                perms = [p.strip() for p in parsed.split(",") if p.strip()]
        else:
            perms = []

        created_at = row[5]
        return {
            "username": row[0],
            "salt": row[1],
            "password_hash": row[2],
            "permissions": perms,
            "is_active": bool(row[4]),
            "created_at": created_at.isoformat()
            if isinstance(created_at, datetime)
            else str(created_at),
        }
    except Exception:
        return None


def verify_user(username: str, password: str) -> bool:
    """
    Kullanıcı adı / parola kontrolü.
    """
    u = find_user(username)
    if not u or not u.get("is_active", True):
        return False
    salt = u.get("salt") or ""
    expected = u.get("password_hash") or ""
    return hash_password(password, salt) == expected


# ============================================================
#  KESİM TİPİ & SÜS KENAR & KISIT LİSTELERİ – SQL
# ============================================================

def load_blocked_looms() -> list[str]:
    """
    Arızalı / bakımda tezgâhları SQL Server'dan okur.
    """
    try:
        with _sql_conn() as c:
            cur = c.cursor()
            cur.execute("SELECT LoomNo FROM dbo.BlockedLooms ORDER BY LoomNo;")
            rows = cur.fetchall()
        return [str(r[0]) for r in rows]
    except Exception:
        return []


def save_blocked_looms(items: list[str]) -> None:
    """
    Verilen tezgâh listesini SQL'e yazar (tamamen yeniler).
    """
    vals = [
        re.findall(r"\d+", str(x))[0]
        for x in (items or [])
        if re.findall(r"\d+", str(x))
    ]
    uniq = sorted(set(vals))
    try:
        with _sql_conn() as c:
            cur = c.cursor()
            cur.execute("DELETE FROM dbo.BlockedLooms;")
            for loom in uniq:
                cur.execute(
                    "INSERT INTO dbo.BlockedLooms (LoomNo) VALUES (?);",
                    (loom,),
                )
            c.commit()
    except Exception:
        pass


def load_dummy_looms() -> list[str]:
    """
    Boş / dummy gösterilecek tezgâhları SQL Server'dan okur.
    """
    try:
        with _sql_conn() as c:
            cur = c.cursor()
            cur.execute("SELECT LoomNo FROM dbo.DummyLooms ORDER BY LoomNo;")
            rows = cur.fetchall()
        return [str(r[0]) for r in rows]
    except Exception:
        return []


def save_dummy_looms(items: list[str]) -> None:
    """
    Verilen dummy tezgâh listesini SQL'e yazar (tamamen yeniler).
    """
    vals = [
        re.findall(r"\d+", str(x))[0]
        for x in (items or [])
        if re.findall(r"\d+", str(x))
    ]
    uniq = sorted(set(vals))
    try:
        with _sql_conn() as c:
            cur = c.cursor()
            cur.execute("DELETE FROM dbo.DummyLooms;")
            for loom in uniq:
                cur.execute(
                    "INSERT INTO dbo.DummyLooms (LoomNo) VALUES (?);",
                    (loom,),
                )
            c.commit()
    except Exception:
        pass


def load_loom_cut_map() -> dict:
    """
    Tezgah -> Kesim Tipi eşlemesini SQL Server'dan okur.
    Dönüş: {"2201": "ISAVER", "2202": "ROTOCUT", ...}
    """
    try:
        with _sql_conn() as c:
            cur = c.cursor()
            cur.execute("SELECT LoomNo, CutType FROM dbo.LoomCutMap;")
            rows = cur.fetchall()
        return {str(r[0]): str(r[1]) for r in rows}
    except Exception:
        return {}


def save_loom_cut_map(d: dict) -> None:
    """
    Verilen dict'i komple SQL'e yazar (öncekileri silip baştan yazar).
    d: {"2201": "ISAVER", "2202": "ROTOCUT", ...}
    """
    if not isinstance(d, dict):
        return

    try:
        with _sql_conn() as c:
            cur = c.cursor()
            cur.execute("DELETE FROM dbo.LoomCutMap;")

            for loom, ctype in d.items():
                loom_str = str(loom).strip()
                cut_str = str(ctype).strip()
                if not loom_str or not cut_str:
                    continue
                cur.execute(
                    "INSERT INTO dbo.LoomCutMap (LoomNo, CutType) VALUES (?, ?);",
                    (loom_str, cut_str),
                )
            c.commit()
    except Exception:
        pass


def load_type_selvedge_map() -> dict:
    """
    Kök Tip -> Süs Kenar eşleşmesini SQL'den okur.
    Dönüş: {"KOK123": "SüsAçıklama", ...}
    """
    try:
        with _sql_conn() as c:
            cur = c.cursor()
            cur.execute("SELECT RootType, Selvedge FROM dbo.TypeSelvedgeMap;")
            rows = cur.fetchall()
        return {str(r[0]): str(r[1]) for r in rows}
    except Exception:
        return {}


def save_type_selvedge_map(d: dict) -> None:
    """
    Verilen dict'i komple SQL'e yazar.
    d: {"KOK123": "SüsAçıklama", ...}
    """
    if not isinstance(d, dict):
        return

    try:
        with _sql_conn() as c:
            cur = c.cursor()
            cur.execute("DELETE FROM dbo.TypeSelvedgeMap;")

            for root, sel in d.items():
                root_str = str(root).strip()
                sel_str = str(sel).strip()
                if not root_str or not sel_str:
                    continue
                cur.execute(
                    "INSERT INTO dbo.TypeSelvedgeMap (RootType, Selvedge) VALUES (?, ?);",
                    (root_str, sel_str),
                )
            c.commit()
    except Exception:
        pass


# ============================================================
#  USTA DEFTERİ – SAYIM FONKSİYONU (SQL)
# ============================================================

def load_usta_dataframe(sqlite_path: str | None = None) -> pd.DataFrame:
    """
    Eski API'yi bozmamak için SQL türevi.
    _ts: datetime, _what: 'DÜĞÜM' / 'TAKIM', _dir: şimdilik boş string.
    """
    try:
        with _sql_conn() as c:
            cur = c.cursor()
            cur.execute("SELECT Id, Tarih, IsTanimi FROM dbo.UstaDefteri;")
            rows = cur.fetchall()
            cols = [d[0] for d in cur.description]
            df = pd.DataFrame.from_records(rows, columns=cols)
    except Exception:
        return pd.DataFrame(columns=["_ts", "_what", "_dir"])

    # Tarih -> datetime
    try:
        ts = pd.to_datetime(df["Tarih"], errors="coerce")
    except Exception:
        ts = pd.NaT

    df["_ts"] = ts
    df["_what"] = df.get("IsTanimi", "").astype(str).str.upper()
    df["_dir"] = ""  # Şimdilik yön bilgisi yok

    return df[["_ts", "_what", "_dir"]]


def count_usta_between(
    start_dt: datetime, end_dt: datetime, what: str = "DÜĞÜM", direction: str = "ALINDI"
) -> int:
    """
    [start_dt, end_dt) aralığında Usta Defteri'nden sayım.
    Şu an yön (ALINDI/VERİLDİ) bilgisi tabloya kaydedilmediği için sadece IsTanimi bazında sayım yapılır.
    """
    try:
        s_date = start_dt.date()
        e_date = end_dt.date()
        w = str(what).upper().strip()

        sql = """
        SELECT COUNT(*)
        FROM dbo.UstaDefteri
        WHERE Tarih >= ? AND Tarih < ?
        """
        params: list[object] = [s_date, e_date]

        if w:
            sql += " AND UPPER(IsTanimi) = ?"
            params.append(w)

        with _sql_conn() as c:
            cur = c.cursor()
            cur.execute(sql, tuple(params))
            row = cur.fetchone()
            if row and row[0] is not None:
                return int(row[0])
    except Exception:
        pass

    return 0

def load_usta_etiket_tezgah_map() -> dict[str, str]:
        """
        UstaDefteri tablosundaki EtiketNo → Tezgah eşleşmesini (en güncel kayıt öncelikli) döndürür.

        İsimler metin olarak normalize edilir; boş/NaN değerler hariç tutulur.
        """

        def _clean(val) -> str:
            if val is None:
                return ""
            try:
                if isinstance(val, float) and pd.isna(val):
                    return ""
            except Exception:
                pass

            s = str(val).strip()
            if not s:
                return ""
            # "1234.0" → "1234"
            s = re.sub(r"\.0+$", "", s)
            return s

        sql = """
        SELECT EtiketNo, Tezgah
        FROM dbo.UstaDefteri
        WHERE EtiketNo IS NOT NULL AND LTRIM(RTRIM(EtiketNo)) <> ''
        ORDER BY Id DESC;
        """

        try:
            with _sql_conn() as c:
                cur = c.cursor()
                cur.execute(sql)
                rows = cur.fetchall()
        except Exception:
            return {}

        mapping: dict[str, str] = {}
        for row in rows:
            etiket = _clean(row[0] if len(row) > 0 else None)
            tezgah = _clean(row[1] if len(row) > 1 else None)
            if not etiket or not tezgah:
                continue
            # Id DESC ile geldiği için ilk denk gelen en güncel; daha eski kayıtları ezme.
            mapping.setdefault(etiket, tezgah)
        return mapping
def fetch_tip_buzulme_model(tip_kodlari: list[str]) -> pd.DataFrame:
    """
    dbo.TipBuzulmeModel tablosundan TipKodu bazında:
      TipKodu, GecmisBuzulme, SistemBuzulme, GuvenAraligi
    döndürür.

    tip_kodlari boşsa boş df döner.
    """
    tips = [str(x).strip() for x in (tip_kodlari or []) if str(x).strip()]
    if not tips:
        return pd.DataFrame(columns=["TipKodu", "GecmisBuzulme", "SistemBuzulme", "GuvenAraligi"])

    # IN param limitlerine takılmamak için parça parça çek
    out_frames: list[pd.DataFrame] = []
    chunk = 800  # güvenli
    sql_tpl = """
    SELECT TipKodu, GecmisBuzulme, SistemBuzulme, GuvenAraligi
    FROM dbo.TipBuzulmeModel
    WHERE TipKodu IN ({})
    """

    try:
        with _sql_conn() as c:
            for i in range(0, len(tips), chunk):
                part = tips[i:i+chunk]
                placeholders = ",".join(["?"] * len(part))
                sql = sql_tpl.format(placeholders)
                import warnings
                warnings.filterwarnings(
                    "ignore",
                    message="pandas only supports SQLAlchemy connectable.*",
                    category=UserWarning,
                )

                df = pd.read_sql(sql, c, params=part)
                out_frames.append(df)
    except Exception:
        return pd.DataFrame(columns=["TipKodu", "GecmisBuzulme", "SistemBuzulme", "GuvenAraligi"])

    if not out_frames:
        return pd.DataFrame(columns=["TipKodu", "GecmisBuzulme", "SistemBuzulme", "GuvenAraligi"])

    res = pd.concat(out_frames, ignore_index=True)
    # TipKodu tekilleştir
    res = res.drop_duplicates(subset=["TipKodu"], keep="last")
    return res

