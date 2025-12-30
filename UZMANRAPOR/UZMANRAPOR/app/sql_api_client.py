from __future__ import annotations

import json
import os
import urllib.error
import urllib.request
from typing import Any, Iterable, Optional


class SqlApiError(RuntimeError):
    pass


def _env(name: str, default: str) -> str:
    value = os.getenv(name)
    return value.strip() if value else default


class ApiConnection:
    """
    pyodbc benzeri minimal bir arayüz sağlayan HTTP tabanlı "bağlantı".
    Amaç: mevcut kod akışını bozmadan SQL erişimini API arkasına almak.
    """

    def __init__(
        self,
        base_url: Optional[str] = None,
        endpoint: Optional[str] = None,
        timeout: int = 30,
        token: Optional[str] = None,
    ) -> None:
        self.base_url = (base_url or _env("UZMANRAPOR_API_URL", "http://10.30.1.68:8000")).rstrip("/")
        raw_endpoint = endpoint or _env("UZMANRAPOR_SQL_ENDPOINT", "/sql")
        self.endpoint = raw_endpoint if raw_endpoint.startswith("/") else f"/{raw_endpoint}"
        self.timeout = timeout
        self.token = token or _env("UZMANRAPOR_API_TOKEN", "")

    def cursor(self) -> "ApiCursor":
        return ApiCursor(self)

    def commit(self) -> None:
        return

    def close(self) -> None:
        return

    def __enter__(self) -> "ApiConnection":
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        self.close()

    def _request(self, payload: dict[str, Any]) -> dict[str, Any]:
        url = f"{self.base_url}{self.endpoint}"
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        headers = {"Content-Type": "application/json"}
        if self.token:
            headers["X-Token"] = self.token
        req = urllib.request.Request(url, data=data, headers=headers, method="POST")
        try:
            with urllib.request.urlopen(req, timeout=self.timeout) as resp:
                body = resp.read().decode("utf-8")
        except urllib.error.HTTPError as exc:
            try:
                msg = exc.read().decode("utf-8")
            except Exception:
                msg = ""
            raise SqlApiError(f"SQL API hatası: {exc.code} {exc.reason} {msg}".strip()) from exc
        except urllib.error.URLError as exc:
            raise SqlApiError(f"SQL API bağlantı hatası: {exc.reason}") from exc

        try:
            data = json.loads(body)
        except json.JSONDecodeError as exc:
            raise SqlApiError("SQL API geçersiz JSON döndürdü.") from exc

        if isinstance(data, dict) and data.get("error"):
            raise SqlApiError(str(data["error"]))
        if not isinstance(data, dict):
            raise SqlApiError("SQL API beklenmeyen cevap döndürdü.")
        return data


class ApiCursor:
    def __init__(self, conn: ApiConnection) -> None:
        self._conn = conn
        self._rows: list[list[Any]] = []
        self.description: Optional[list[tuple[Any, ...]]] = None
        self.rowcount: int = -1

    def execute(self, query: str, params: Optional[Iterable[Any]] = None) -> "ApiCursor":
        q = (query or "").strip()
        if q.endswith(";"):
            q = q[:-1].rstrip()  # sadece sondaki ; temizlensin

        payload = {"query": q, "params": list(params or [])}
        data = self._conn._request(payload)

        rows = data.get("rows")
        columns = data.get("columns")
        if rows is None and "data" in data:
            rows = data.get("data")
        if columns is None and "column_names" in data:
            columns = data.get("column_names")

        if rows is None:
            rows = []

        # dict row support
        if rows and isinstance(rows[0], dict):
            if not columns:
                columns = list(rows[0].keys())
            rows = [[row.get(col) for col in columns] for row in rows]

        self._rows = [list(row) for row in rows]
        if columns:
            self.description = [(col, None, None, None, None, None, None) for col in columns]
        else:
            self.description = None

        self.rowcount = data.get("rowcount")
        if self.rowcount is None:
            self.rowcount = data.get("affected_rows", len(self._rows) if self._rows else -1)
        return self

    def fetchone(self) -> Optional[tuple[Any, ...]]:
        if not self._rows:
            return None
        row = self._rows.pop(0)
        return tuple(row)

    def fetchall(self) -> list[tuple[Any, ...]]:
        rows = [tuple(row) for row in self._rows]
        self._rows = []
        return rows


def get_sql_connection() -> ApiConnection:
    return ApiConnection()
