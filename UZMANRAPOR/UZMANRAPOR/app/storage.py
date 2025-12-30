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
import zlib

import pandas as pd
from app.sql_api_client import ApiConnection, get_sql_connection
from app.db_name import DB_NAME



# ============================================================
#  SQL SERVER BAĞLANTISI (FastAPI üzerinden)
# ============================================================

def _sql_conn() -> ApiConnection:
    """UzmanRaporDB bağlantısı (FastAPI üzerinden)."""
    return get_sql_connection()


def _fetch_dataframe(sql: str, params: tuple | list | None = None) -> pd.DataFrame:
    """API cursor ile DataFrame üretir (pd.read_sql yerine)."""
    with _sql_conn() as c:
        cur = c.cursor()
        cur.execute(sql, params or [])
        rows = cur.fetchall()
        columns = [d[0] for d in cur.description] if cur.description else []
    return pd.DataFrame(rows, columns=columns)


# ============================================================
#  APP META (GENEL ANAHTAR/DEĞER)
#  Not: API modunda DDL yok. Tablolar SSMS'de yönetilecek.
# ============================================================

def _ensure_meta_table() -> None:
    return


def _meta_get(key: str) -> str | None:
    _ensure_meta_table()
    try:
        with _sql_conn() as c:
            cur = c.cursor()
            cur.execute(f"SELECT MetaValue FROM [{DB_NAME}].[dbo].[AppMeta] WHERE MetaKey = ?;", (key,))
            row = cur.fetchone()
        if not row:
            return None
        return row[0]
    except Exception:
        return None


def _meta_set(key: str, value: str | None) -> None:
    _ensure_meta_table()
    try:
        with _sql_conn() as c:
            cur = c.cursor()

            # 1) Önce UPDATE dene
            cur.execute(
                f"UPDATE [{DB_NAME}].[dbo].[AppMeta] "
                "SET MetaValue = ?, UpdatedAt = SYSUTCDATETIME() "
                "WHERE MetaKey = ?",
                (value, key),
            )

            # 2) Rowcount API’de güvenilir olmayabilir; var mı diye kontrol et
            cur.execute(
                f"SELECT COUNT(*) FROM [{DB_NAME}].[dbo].[AppMeta] WHERE MetaKey = ?",
                (key,),
            )
            cnt = cur.fetchone()[0] if cur.fetchone is not None else 0

            if cnt == 0:
                cur.execute(
                    f"INSERT INTO [{DB_NAME}].[dbo].[AppMeta] (MetaKey, MetaValue, UpdatedAt) "
                    "VALUES (?, ?, SYSUTCDATETIME())",
                    (key, value),
                )

            c.commit()
    except Exception as e:
        print(f"[APPMETA] yazma hatası: {e!r}")



# ============================================================
#  NOT KURALLARI (SQL)
# ============================================================

def _note_rules_table_exists() -> bool:
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
        if isinstance(raw, bytes):
            raw = raw.decode("utf-8", errors="ignore")
        blob = base64.b64decode(raw)
        obj = pickle.loads(blob)
        if isinstance(obj, list):
            return [r for r in obj if isinstance(r, dict)]
    except Exception:
        pass
    return []


def load_rules() -> list[dict]:
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
                # API tarafında RuleData nvarchar(max) base64 string tutuluyorsa:
                if isinstance(blob, str):
                    blob = base64.b64decode(blob)
                rule = pickle.loads(blob)
                if isinstance(rule, dict):
                    rules.append(rule)
            except Exception:
                continue
    except Exception:
        pass

    return rules


def save_rules(rules: list[dict]) -> None:
    if not isinstance(rules, list):
        return

    cleaned = [r for r in rules if isinstance(r, dict)]

    # AppMeta'ya base64 pickle olarak yaz
    try:
        if cleaned:
            blob = pickle.dumps(cleaned)
            payload = base64.b64encode(blob).decode("ascii")
            _meta_set("note_rules", payload)
        else:
            _meta_set("note_rules", None)
    except Exception:
        pass

    # NoteRules tablosu varsa orayı da güncelle
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
                payload_blob = base64.b64encode(blob).decode("ascii")
                cur.execute(
                    "INSERT INTO dbo.NoteRules (RuleData) VALUES (?);",
                    (payload_blob,),
                )
            c.commit()
    except Exception:
        pass


# ============================================================
#  SON GÜNCELLEME
# ============================================================

def load_last_update() -> datetime | None:
    raw = _meta_get("last_update")
    if not raw:
        return None
    try:
        dt = datetime.fromisoformat(raw)
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=ZoneInfo("Europe/Istanbul"))
        return dt
    except Exception:
        return None


def save_last_update(dt: datetime) -> None:
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=ZoneInfo("Europe/Istanbul"))
    _meta_set("last_update", dt.isoformat())


# ============================================================
#  SNAPSHOTS
#  Not: API modunda DDL yok. [{DB_NAME}].[dbo].[Snapshots] SSMS’de hazır olmalı.
# ============================================================

def _ensure_snapshot_table() -> None:
    return


def save_df_snapshot(df: pd.DataFrame | None, which: str) -> None:
    if df is None:
        return

    _ensure_snapshot_table()

    try:
        buf = io.BytesIO()
        df.to_pickle(buf)
        raw_bytes = buf.getvalue()

        compressed = zlib.compress(raw_bytes, level=9)
        hex_str = compressed.hex()

        with _sql_conn() as c:
            cur = c.cursor()
            cur.execute(f"DELETE FROM [{DB_NAME}].[dbo].[Snapshots] WHERE Name = ?;", (which,))
            cur.execute(
                f"INSERT INTO [{DB_NAME}].[dbo].[Snapshots] (Name, DataHex) VALUES (?, ?);",
                (which, hex_str),
            )
            c.commit()
    except Exception as e:
        print(f"[SNAPSHOT] {which}: KAYIT HATASI -> {e!r}")


def load_df_snapshot(which: str) -> pd.DataFrame | None:
    _ensure_snapshot_table()

    try:
        with _sql_conn() as c:
            cur = c.cursor()
            cur.execute(f"SELECT DataHex FROM [{DB_NAME}].[dbo].[Snapshots] WHERE Name = ?;", (which,))
            row = cur.fetchone()

        if not row or row[0] is None:
            return None

        compressed = bytes.fromhex(row[0])
        raw_bytes = zlib.decompress(compressed)
        buf = io.BytesIO(raw_bytes)

        df = pd.read_pickle(buf)
        return df if isinstance(df, pd.DataFrame) else None
    except Exception as e:
        print(f"[SNAPSHOT] {which}: YÜKLEME HATASI -> {e!r}")
        return None


# ============================================================
#  KULLANICI VARSAYILANI
# ============================================================

def get_username_default() -> str:
    val = _meta_get("last_username")
    return str(val) if val else "Anonim"


def set_username_default(name: str) -> None:
    name = (name or "").strip()
    if not name:
        return
    _meta_set("last_username", name)


# ============================================================
#  LOGIN SİSTEMİ – AppUsers (SQL)
#   Not: API modunda DDL yok. [{DB_NAME}].[dbo].[AppUsers] SSMS’de hazır olmalı.
# ============================================================

def hash_password(password: str, salt: str) -> str:
    payload = f"{salt}:{password}".encode("utf-8")
    return hashlib.sha256(payload).hexdigest()


def _ensure_app_users_table() -> None:
    return


def ensure_user_db() -> None:
    """
    Tablo yoksa sessizce çıkar.
    Varsa boşsa default admin eklemeyi dener.
    """
    _ensure_app_users_table()
    try:
        with _sql_conn() as c:
            cur = c.cursor()
            cur.execute(f"SELECT COUNT(*) FROM [{DB_NAME}].[dbo].[AppUsers];")
            row = cur.fetchone()
            count = int(row[0]) if row and row[0] is not None else 0

            if count == 0:
                salt = secrets.token_hex(16)
                pwd_hash = hash_password("admin", salt)
                perms = "admin,read,write"
                cur.execute(
                    f"INSERT INTO [{DB_NAME}].[dbo].[AppUsers] (Username, Salt, PasswordHash, Permissions, IsActive) "
                    "VALUES (?, ?, ?, ?, 1);",
                    ("admin", salt, pwd_hash, perms),
                )
            c.commit()
    except Exception:
        pass


def load_users() -> list[dict]:
    ensure_user_db()
    users: list[dict] = []

    try:
        with _sql_conn() as c:
            cur = c.cursor()
            cur.execute(
                f"SELECT Username, Salt, PasswordHash, Permissions, IsActive, CreatedAt "
                f"FROM [{DB_NAME}].[dbo].[AppUsers];"
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
                try:
                    tmp = ast.literal_eval(perms_raw)
                    parsed = tmp if isinstance(tmp, list) else perms_raw
                except Exception:
                    parsed = perms_raw
                if isinstance(parsed, list):
                    perms = [str(p).strip() for p in parsed if str(p).strip()]
                else:
                    perms = [p.strip() for p in str(parsed).split(",") if p.strip()]
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
    except Exception as e:
        print("[load_users] ERROR:", e)

    return users


def save_users(users: list[dict]) -> None:
    ensure_user_db()
    if not isinstance(users, list):
        return

    try:
        with _sql_conn() as c:
            cur = c.cursor()
            cur.execute(f"DELETE FROM [{DB_NAME}].[dbo].[AppUsers];")

            for u in users:
                username = str(u.get("username", "")).strip()
                salt = str(u.get("salt", "")).strip()
                pwd_hash = str(u.get("password_hash", "")).strip()
                perms = u.get("permissions", [])

                if not username or not salt or not pwd_hash:
                    continue

                perms_raw = ",".join([str(p).strip() for p in perms]) if isinstance(perms, list) else str(perms)
                is_active_bit = 1 if u.get("is_active", True) else 0

                cur.execute(
                    f"INSERT INTO [{DB_NAME}].[dbo].[AppUsers] (Username, Salt, PasswordHash, Permissions, IsActive) "
                    "VALUES (?, ?, ?, ?, ?);",
                    (username, salt, pwd_hash, perms_raw, is_active_bit),
                )

            c.commit()
    except Exception:
        pass


def find_user(username: str) -> dict | None:
    ensure_user_db()
    try:
        with _sql_conn() as c:
            cur = c.cursor()
            cur.execute(
                f"SELECT Username, Salt, PasswordHash, Permissions, IsActive, CreatedAt "
                f"FROM [{DB_NAME}].[dbo].[AppUsers] WHERE Username = ?;",
                (username,),
            )
            row = cur.fetchone()

        if not row:
            return None

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
                perms = [p.strip() for p in str(parsed).split(",") if p.strip()]
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
    u = find_user(username)
    if not u or not u.get("is_active", True):
        return False
    salt = u.get("salt") or ""
    expected = u.get("password_hash") or ""
    return hash_password(password, salt) == expected


# ============================================================
#  BLOK/DUMMY/CUT/SELVEDGE MAP
# ============================================================

def load_blocked_looms() -> list[str]:
    try:
        with _sql_conn() as c:
            cur = c.cursor()
            cur.execute(f"SELECT LoomNo FROM [{DB_NAME}].[dbo].[BlockedLooms] ORDER BY LoomNo;")
            rows = cur.fetchall()
        return [str(r[0]) for r in rows]
    except Exception:
        return []


def save_blocked_looms(items: list[str]) -> None:
    vals = [re.findall(r"\d+", str(x))[0] for x in (items or []) if re.findall(r"\d+", str(x))]
    uniq = sorted(set(vals))
    try:
        with _sql_conn() as c:
            cur = c.cursor()
            cur.execute(f"DELETE FROM [{DB_NAME}].[dbo].[BlockedLooms];")
            for loom in uniq:
                cur.execute(f"INSERT INTO [{DB_NAME}].[dbo].[BlockedLooms] (LoomNo) VALUES (?);", (loom,))
            c.commit()
    except Exception:
        pass


def load_dummy_looms() -> list[str]:
    try:
        with _sql_conn() as c:
            cur = c.cursor()
            cur.execute(f"SELECT LoomNo FROM [{DB_NAME}].[dbo].[DummyLooms] ORDER BY LoomNo;")
            rows = cur.fetchall()
        return [str(r[0]) for r in rows]
    except Exception:
        return []


def save_dummy_looms(items: list[str]) -> None:
    vals = [re.findall(r"\d+", str(x))[0] for x in (items or []) if re.findall(r"\d+", str(x))]
    uniq = sorted(set(vals))
    try:
        with _sql_conn() as c:
            cur = c.cursor()
            cur.execute(f"DELETE FROM [{DB_NAME}].[dbo].[DummyLooms];")
            for loom in uniq:
                cur.execute(f"INSERT INTO [{DB_NAME}].[dbo].[DummyLooms] (LoomNo) VALUES (?);", (loom,))
            c.commit()
    except Exception:
        pass


def load_loom_cut_map() -> dict:
    try:
        with _sql_conn() as c:
            cur = c.cursor()
            cur.execute(f"SELECT LoomNo, CutType FROM [{DB_NAME}].[dbo].[LoomCutMap];")
            rows = cur.fetchall()
        return {str(r[0]): str(r[1]) for r in rows}
    except Exception:
        return {}


def save_loom_cut_map(d: dict) -> None:
    if not isinstance(d, dict):
        return
    try:
        with _sql_conn() as c:
            cur = c.cursor()
            cur.execute(f"DELETE FROM [{DB_NAME}].[dbo].[LoomCutMap];")
            for loom, ctype in d.items():
                loom_str = str(loom).strip()
                cut_str = str(ctype).strip()
                if loom_str and cut_str:
                    cur.execute(f"INSERT INTO [{DB_NAME}].[dbo].[LoomCutMap] (LoomNo, CutType) VALUES (?, ?);", (loom_str, cut_str))
            c.commit()
    except Exception:
        pass


def load_type_selvedge_map() -> dict:
    try:
        with _sql_conn() as c:
            cur = c.cursor()
            cur.execute(f"SELECT RootType, Selvedge FROM [{DB_NAME}].[dbo].[TypeSelvedgeMap];")
            rows = cur.fetchall()
        return {str(r[0]): str(r[1]) for r in rows}
    except Exception:
        return {}


def save_type_selvedge_map(d: dict) -> None:
    if not isinstance(d, dict):
        return
    try:
        with _sql_conn() as c:
            cur = c.cursor()

            for root, sel in d.items():
                root_str = str(root).strip().upper()
                sel_str = str(sel).strip()
                if not root_str or not sel_str:
                    continue

                # UPDATE dene
                cur.execute(
                    f"UPDATE [{DB_NAME}].[dbo].[TypeSelvedgeMap] "
                    "SET Selvedge = ? "
                    "WHERE RootType = ?",
                    (sel_str, root_str),
                )

                # Var mı kontrol et; yoksa INSERT
                cur.execute(
                    f"SELECT COUNT(*) FROM [{DB_NAME}].[dbo].[TypeSelvedgeMap] WHERE RootType = ?",
                    (root_str,),
                )
                cnt = cur.fetchone()[0]
                if cnt == 0:
                    cur.execute(
                        f"INSERT INTO [{DB_NAME}].[dbo].[TypeSelvedgeMap] (RootType, Selvedge) VALUES (?, ?)",
                        (root_str, sel_str),
                    )

            c.commit()
    except Exception as e:
        print(f"[TypeSelvedgeMap] yazma hatası: {e!r}")



# ============================================================
#  USTA DEFTERİ – SAYIM (SQL)
# ============================================================

def load_usta_dataframe(sqlite_path: str | None = None) -> pd.DataFrame:
    try:
        with _sql_conn() as c:
            cur = c.cursor()
            cur.execute(f"SELECT Id, Tarih, IsTanimi FROM [{DB_NAME}].[dbo].[UstaDefteri];")
            rows = cur.fetchall()
            cols = [d[0] for d in cur.description] if cur.description else ["Id", "Tarih", "IsTanimi"]
            df = pd.DataFrame.from_records(rows, columns=cols)
    except Exception:
        return pd.DataFrame(columns=["_ts", "_what", "_dir"])

    try:
        ts = pd.to_datetime(df["Tarih"], errors="coerce")
    except Exception:
        ts = pd.NaT

    df["_ts"] = ts
    df["_what"] = df.get("IsTanimi", "").astype(str).str.upper()
    df["_dir"] = ""
    return df[["_ts", "_what", "_dir"]]


def count_usta_between(start_dt: datetime, end_dt: datetime, what: str = "DÜĞÜM", direction: str = "ALINDI") -> int:
    try:
        s_date = start_dt.date()
        e_date = end_dt.date()
        w = str(what).upper().strip()

        sql = f"SELECT COUNT(*) FROM [{DB_NAME}].[dbo].[UstaDefteri] WHERE Tarih >= ? AND Tarih < ?"
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
        s = re.sub(r"\.0+$", "", s)
        return s

    sql = f"""
    SELECT EtiketNo, Tezgah
    FROM [{DB_NAME}].[dbo].[UstaDefteri]
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
        if etiket and tezgah:
            mapping.setdefault(etiket, tezgah)
    return mapping


def fetch_tip_buzulme_model(tip_kodlari: list[str]) -> pd.DataFrame:
    tips = [str(x).strip() for x in (tip_kodlari or []) if str(x).strip()]
    if not tips:
        return pd.DataFrame(columns=["TipKodu", "GecmisBuzulme", "SistemBuzulme", "GuvenAraligi"])

    out_frames: list[pd.DataFrame] = []
    chunk = 800
    sql_tpl = f"""
    SELECT TipKodu, GecmisBuzulme, SistemBuzulme, GuvenAraligi
    FROM [{DB_NAME}].[dbo].[TipBuzulmeModel]
    WHERE TipKodu IN ({{}})
    """

    try:
        for i in range(0, len(tips), chunk):
            part = tips[i:i + chunk]
            placeholders = ",".join(["?"] * len(part))
            sql = sql_tpl.format(placeholders)
            df = _fetch_dataframe(sql, part)
            out_frames.append(df)
    except Exception:
        return pd.DataFrame(columns=["TipKodu", "GecmisBuzulme", "SistemBuzulme", "GuvenAraligi"])

    if not out_frames:
        return pd.DataFrame(columns=["TipKodu", "GecmisBuzulme", "SistemBuzulme", "GuvenAraligi"])

    res = pd.concat(out_frames, ignore_index=True)
    res = res.drop_duplicates(subset=["TipKodu"], keep="last")
    return res
