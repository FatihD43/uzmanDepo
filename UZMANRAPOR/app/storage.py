from __future__ import annotations
from pathlib import Path
from datetime import datetime
from zoneinfo import ZoneInfo

import os
import json
import re
import secrets
import hashlib
import io

import pandas as pd
import pyodbc

# ============================================================
#  GENEL YOLLAR (Kişisel ayarlar – her PC için)
# ============================================================

APP_DIR = Path.home() / ".uzman_rapor"
APP_DIR.mkdir(parents=True, exist_ok=True)

RULES_PATH = APP_DIR / "notes_rules.json"
META_PATH = APP_DIR / "meta.json"
USERCFG_PATH = APP_DIR / "user.json"

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


# ============================================================
#  NOT KURALLARI (kalıcı, ama yerel JSON'da kalsın)
# ============================================================

def load_rules() -> list[dict]:
    if RULES_PATH.exists():
        try:
            with open(RULES_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, list):
                return data
        except Exception:
            pass
    return []


def save_rules(rules: list[dict]) -> None:
    try:
        with open(RULES_PATH, "w", encoding="utf-8") as f:
            json.dump(rules or [], f, ensure_ascii=False, indent=2)
    except Exception:
        pass


# ============================================================
#  SON GÜNCELLEME (Planlama tıklanınca kaydedilen zaman)
# ============================================================

def load_last_update() -> datetime | None:
    if META_PATH.exists():
        try:
            with open(META_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
            iso = data.get("last_update_iso")
            if iso:
                dt = datetime.fromisoformat(iso)
                # naive ise İstanbul TZ ata
                if dt.tzinfo is None:
                    dt = dt.replace(tzinfo=ZoneInfo("Europe/Istanbul"))
                return dt
        except Exception:
            pass
    return None


def save_last_update(dt: datetime) -> None:
    try:
        # dt naive ise İstanbul TZ'li kabul et
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=ZoneInfo("Europe/Istanbul"))
        meta = {"last_update_iso": dt.isoformat()}
        if META_PATH.exists():
            try:
                with open(META_PATH, "r", encoding="utf-8") as f:
                    cur = json.load(f)
                if isinstance(cur, dict):
                    cur.update(meta)
                    meta = cur
            except Exception:
                pass
        with open(META_PATH, "w", encoding="utf-8") as f:
            json.dump(meta, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


# ============================================================
#  SNAPSHOT (Dinamik & Running) – SQL TABLOSU
# ============================================================

def _ensure_snapshot_table() -> None:
    """
    Snapshots tablosu:
      Name: 'dinamik' veya 'running' vb
      Data: varbinary(max) (pandas pickle)
    """
    sql = """
    IF OBJECT_ID('dbo.Snapshots', 'U') IS NULL
    BEGIN
        CREATE TABLE dbo.Snapshots (
            Name      nvarchar(50) NOT NULL PRIMARY KEY,
            Data      varbinary(max) NULL,
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
        pass


def save_df_snapshot(df: pd.DataFrame | None, which: str) -> None:
    """
    Dinamik / Running DataFrame'lerini SQL'e pickle olarak yazar.
    which: 'dinamik' veya 'running' vb.
    """
    if df is None:
        return

    _ensure_snapshot_table()

    try:
        buf = io.BytesIO()
        df.to_pickle(buf)
        data_bytes = buf.getvalue()

        with _sql_conn() as c:
            cur = c.cursor()
            # Önce var olan kaydı sil
            cur.execute("DELETE FROM dbo.Snapshots WHERE Name = ?;", (which,))
            # Sonra yeni kaydı ekle
            cur.execute(
                "INSERT INTO dbo.Snapshots (Name, Data) VALUES (?, ?);",
                (which, pyodbc.Binary(data_bytes)),
            )
            c.commit()
    except Exception:
        pass


def load_df_snapshot(which: str) -> pd.DataFrame | None:
    """
    Dinamik / Running snapshot'ı SQL'den okur.
    Kayıt yoksa None döner.
    """
    _ensure_snapshot_table()

    try:
        with _sql_conn() as c:
            cur = c.cursor()
            cur.execute("SELECT Data FROM dbo.Snapshots WHERE Name = ?;", (which,))
            row = cur.fetchone()

        if not row or row[0] is None:
            return None

        buf = io.BytesIO(row[0])
        df = pd.read_pickle(buf)
        if isinstance(df, pd.DataFrame):
            return df
    except Exception:
        pass

    return None


# ============================================================
#  KULLANICI VARSAYILANI (Sadece bu PC için – yerel JSON)
# ============================================================

def get_username_default() -> str:
    if USERCFG_PATH.exists():
        try:
            with open(USERCFG_PATH, "r", encoding="utf-8") as f:
                d = json.load(f)
            u = d.get("username")
            if u:
                return str(u)
        except Exception:
            pass
    return "Anonim"


def set_username_default(name: str) -> None:
    try:
        d = {"username": name}
        if USERCFG_PATH.exists():
            try:
                with open(USERCFG_PATH, "r", encoding="utf-8") as f:
                    cur = json.load(f)
                if isinstance(cur, dict):
                    cur.update(d)
                    d = cur
            except Exception:
                pass
        with open(USERCFG_PATH, "w", encoding="utf-8") as f:
            json.dump(d, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


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
            Permissions  nvarchar(max) NULL,      -- JSON (örn: ["admin","read","write"])
            IsActive     bit NOT NULL CONSTRAINT DF_AppUsers_IsActive DEFAULT (1),
            CreatedAt    datetime2 NOT NULL CONSTRAINT DF_AppUsers_CreatedAt DEFAULT (sysdatetime())
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
                # Varsayılan admin
                salt = secrets.token_hex(16)
                pwd_hash = hash_password("admin", salt)
                perms = json.dumps(["admin", "read", "write"], ensure_ascii=False)

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

            # permissions JSON mu string mi?
            perms: list[str] | str
            if isinstance(perms_raw, str):
                try:
                    tmp = json.loads(perms_raw)
                    if isinstance(tmp, list):
                        perms = tmp
                    else:
                        perms = perms_raw
                except Exception:
                    perms = perms_raw
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
                    perms_raw = json.dumps(perms, ensure_ascii=False)
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
        if isinstance(perms_raw, str):
            try:
                tmp = json.loads(perms_raw)
                perms = tmp if isinstance(tmp, list) else perms_raw
            except Exception:
                perms = perms_raw
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
    Eski API'yi bozmamak için basit bir SQL türevi.
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
