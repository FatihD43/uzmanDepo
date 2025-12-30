from __future__ import annotations

import base64
import os
import re
from typing import Any

import pyodbc
from fastapi import FastAPI, Header, HTTPException
from pydantic import BaseModel, Field

app = FastAPI(title="UzmanRapor API", version="1.0")


class SqlRequest(BaseModel):
    query: str = Field(..., description="Parametreli SQL (?), veya EXEC dbo.sp_X @p=?")
    params: list[Any] = Field(default_factory=list)


def _env(name: str, default: str = "") -> str:
    v = os.getenv(name)
    return v.strip() if v else default


MAX_ROWS = int(_env("UZMANRAPOR_MAX_ROWS", "20000"))


def _require_token(x_token: str | None) -> None:
    expected = _env("UZMANRAPOR_API_TOKEN", "").strip()

    # Prod davranışı: token config edilmemişse servis yanlış konfigürasyondur
    if not expected:
        raise HTTPException(status_code=500, detail="Server token is not configured")

    if not x_token or x_token != expected:
        raise HTTPException(status_code=401, detail="Unauthorized")


def _sql_conn_str() -> str:
    # Tek satır connection string (tercih)
    raw = _env("UZMANRAPOR_SQL_CONN_STR", "")
    if raw:
        return raw

    # Parça parça (alternatif)
    driver = _env("UZMANRAPOR_SQL_DRIVER", "{ODBC Driver 18 for SQL Server}")
    server = _env("UZMANRAPOR_SQL_SERVER", r"localhost\SQLEXPRESS")
    database = _env("UZMANRAPOR_SQL_DATABASE", "UzmanRaporDB_ISKO14")
    uid = _env("UZMANRAPOR_SQL_UID", "")
    pwd = _env("UZMANRAPOR_SQL_PWD", "")
    trusted = _env("UZMANRAPOR_SQL_TRUSTED", "")

    parts = [f"Driver={driver};", f"Server={server};", f"Database={database};"]
    if trusted.lower() in {"1", "true", "yes"}:
        parts.append("Trusted_Connection=yes;")
    elif uid and pwd:
        parts.append(f"UID={uid};")
        parts.append(f"PWD={pwd};")
    else:
        # Varsayılan: trusted dene
        parts.append("Trusted_Connection=yes;")

    # ODBC 18'de sertifika/şifreleme kaynaklı sorunları engellemek için
    parts.append("TrustServerCertificate=yes;")
    parts.append("Encrypt=no;")

    return "".join(parts)


_FORBIDDEN = re.compile(
    r"\b("
    r"create|alter|drop|truncate|grant|revoke|"
    r"merge|if|begin|end|declare|while|"
    r"use|go|"
    r"xp_|sp_configure|openrowset|openquery"
    r")\b",
    re.IGNORECASE,
)

def _split_csv_env(name: str) -> list[str]:
    raw = (os.getenv(name) or "").strip()
    if not raw:
        return []
    return [x.strip() for x in raw.split(",") if x.strip()]


# 1) Önce çoklu DB env'inden oku (tercih)
_DB_NAMES = _split_csv_env("UZMANRAPOR_SQL_DATABASES")

# 2) Yoksa eski tekli env ile geriye dönük uyum
if not _DB_NAMES:
    one = (os.getenv("UZMANRAPOR_SQL_DATABASE", "UzmanRaporDB") or "").strip() or "UzmanRaporDB"
    _DB_NAMES = [one]


_ALLOWED_TABLES = {
    "AppLookupValues",
    "AppMeta",
    "AppUsers",
    "BlockedLooms",
    "DummyLooms",
    "ItemaAyar",
    "LoomCutMap",
    "Makine_Ayar_Tablosu",
    "NoteRules",
    "Snapshots",
    "TipBuzulmeModel",
    "TypeSelvedgeMap",
    "UstaDefteri",
}

_ALLOWED_OBJECTS = {f"dbo.{t}" for t in _ALLOWED_TABLES}
for _db in _DB_NAMES:
    _ALLOWED_OBJECTS |= {f"{_db}.dbo.{t}" for t in _ALLOWED_TABLES}


_ALLOWED_PROCS = {
    "dbo.sp_ItemaOtomatikAyar",
    "dbo.sp_ItemaTipOzelAyar",
}
for _db in _DB_NAMES:
    _ALLOWED_PROCS.add(f"{_db}.dbo.sp_ItemaOtomatikAyar")
    _ALLOWED_PROCS.add(f"{_db}.dbo.sp_ItemaTipOzelAyar")


# yakalanacak obje referansları: dbo.X veya [db].[dbo].[X]
_OBJ_REF = re.compile(
    r"\b(?:(?:\[?(?P<db>\w+)\]?\.)?\[?dbo\]?\.)\[?(?P<table>\w+)\]?\b",
    re.IGNORECASE,
)
# exec dbo.sp_x veya exec [db].[dbo].[sp_x]
_EXEC_REF = re.compile(
    r"\bexec\s+(?:(?P<db>\[?\w+\]?)\.)?\[?dbo\]?\.\[?(?P<proc>\w+)\]?\b",
    re.IGNORECASE,
)


def _validate_query(query: str) -> None:
    q = query.strip()
    if not q:
        raise HTTPException(status_code=400, detail="Empty query")

    if q.endswith(";"):
        q = q[:-1].rstrip()

    if _FORBIDDEN.search(q):
        raise HTTPException(status_code=403, detail="Forbidden SQL keyword")

    if ";" in q:
        raise HTTPException(status_code=403, detail="Multiple statements are not allowed")

    # izinli komutlar
    head = q.split(None, 1)[0].lower()
    if head not in {"select", "insert", "update", "delete", "exec", "with"}:
        raise HTTPException(status_code=403, detail="Only SELECT/INSERT/UPDATE/DELETE/EXEC are allowed")

    # obje whitelist kontrolü
    objs: set[str] = set()
    for m in _OBJ_REF.finditer(q):
        table = m.group("table")
        db = m.group("db")
        if db:
            objs.add(f"{db}.dbo.{table}")
        else:
            objs.add(f"dbo.{table}")

    for obj in objs:
        if obj not in _ALLOWED_OBJECTS and obj not in _ALLOWED_PROCS:
            raise HTTPException(status_code=403, detail=f"Object not allowed: {obj}")

    if head == "exec":
        m = _EXEC_REF.search(q)
        if not m:
            raise HTTPException(status_code=403, detail="EXEC only allowed for dbo stored procedures")

        db = m.group("db")
        proc = m.group("proc")
        if db:
            db_name = db.strip("[]")
            qualified = f"{db_name}.dbo.{proc}"
            if qualified not in _ALLOWED_PROCS:
                raise HTTPException(status_code=403, detail=f"Procedure not allowed: {qualified}")
        else:
            unqualified = f"dbo.{proc}"
            if unqualified not in _ALLOWED_PROCS:
                raise HTTPException(status_code=403, detail=f"Procedure not allowed: {unqualified}")


def _adapt_params(query: str, params: list[Any]) -> list[Any]:
    # NoteRules varbinary için base64 -> bytes (heuristic)
    q = query.lower()
    if "insert into dbo.noterules" in q and params:
        out = params[:]
        for i, p in enumerate(out):
            if isinstance(p, str):
                try:
                    out[i] = base64.b64decode(p)
                except Exception:
                    pass
        return out
    return params


def _encode_value(v: Any) -> Any:
    if isinstance(v, (bytes, bytearray, memoryview)):
        return base64.b64encode(bytes(v)).decode("ascii")
    return v


@app.get("/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


@app.post("/sql")
def sql(req: SqlRequest, x_token: str | None = Header(default=None)) -> dict[str, Any]:
    conn_str = _sql_conn_str()
    params = _adapt_params(req.query, list(req.params or []))

    conn = None
    try:
        _require_token(x_token)
        _validate_query(req.query)

        conn = pyodbc.connect(conn_str, timeout=10)
        cur = conn.cursor()
        cur.execute(req.query, params)

        if cur.description:
            cols = [d[0] for d in cur.description]
            rows = cur.fetchmany(MAX_ROWS + 1)
            if len(rows) > MAX_ROWS:
                raise HTTPException(
                    status_code=413,
                    detail=f"Result too large (>{MAX_ROWS} rows). Please add filters.",
                )
            data_rows = [[_encode_value(v) for v in row] for row in rows]
            return {"columns": cols, "rows": data_rows, "rowcount": len(data_rows)}

        conn.commit()
        rc = cur.rowcount if cur.rowcount is not None else -1
        return {"columns": [], "rows": [], "affected_rows": rc}

    except HTTPException as e:
        if e.status_code == 403:
            print("[403 FORBIDDEN SQL]", req.query.strip().replace("\n", " ")[:200])
            print("[403 DETAIL]", e.detail)
        raise

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

    finally:
        try:
            if conn is not None:
                conn.close()
        except Exception:
            pass

