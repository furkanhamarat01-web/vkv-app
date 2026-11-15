"""Masaüstü arayüzünün veri katmanı. Kayıtlar HTTP API'de durur."""

import json
import os
import sqlite3
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

OWNED = {"invoices", "reminders", "reports", "tasks", "contacts", "transactions", "projects", "offers"}


class ApiError(RuntimeError):
    pass


class Record(dict):
    def keys(self):
        return list(super().keys())


class _Conn:
    def rollback(self):
        return None


def _token():
    env = os.environ.get("API_TOKEN", "").strip()
    if env:
        return env
    path = Path.home() / "Library" / "Application Support" / "IsMerkezi" / "api.token"
    try:
        return path.read_text(encoding="utf-8").strip()
    except OSError:
        return ""


def _base():
    return os.environ.get(
        "IS_MERKEZI_API",
        "http://hqb6bmzszcsc4t470l10npd5.185.184.210.52.sslip.io",
    ).rstrip("/")


class ApiStore:
    def __init__(self):
        self.base = _base()
        self.token = _token()
        self.owner_id = None
        self._lock_owned = False
        self._workspace_ready = False
        self.conn = _Conn()
        if not self.token:
            raise ApiError("API anahtarı yok. Uygulamayı İş Merkezi simgesinden açın.")
        try:
            self._request("GET", "/health", auth=False)
        except ApiError as exc:
            raise ApiError(f"Sunucuya bağlanılamadı ({self.base}).\n{exc}") from exc

    def close(self):
        return None

    def owns_records(self):
        return self.owner_id is not None or not self._lock_owned

    def set_owner(self, owner_id, claim=False, lock=False, is_admin=False):
        self.owner_id = int(owner_id) if owner_id else None
        self._lock_owned = bool(lock) and self.owner_id is None
        if claim and self.owner_id and is_admin:
            self._request("POST", "/api/store/claim", {"is_admin": True})
        if self.owns_records():
            self.ensure_workspace()

    def ensure_workspace(self):
        if self._lock_owned:
            return
        self._request("POST", "/api/store/workspace", {})
        self._workspace_ready = True

    def all(self, table):
        if table in OWNED and self._lock_owned:
            return []
        return [self._record(item) for item in self._pages(f"/api/store/{table}")]

    def get(self, table, row_id):
        if table in OWNED and self._lock_owned:
            return None
        try:
            return self._record(self._request("GET", f"/api/store/{table}/{int(row_id)}"))
        except ApiError as exc:
            if getattr(exc, "status", None) == 404:
                return None
            raise

    def save(self, table, data, row_id=None):
        if table in OWNED and self._lock_owned:
            raise ValueError("Önce giriş yapılmalı.")
        payload = dict(data)
        if row_id:
            saved = self._request("PATCH", f"/api/store/{table}/{int(row_id)}", payload)
        else:
            saved = self._request("POST", f"/api/store/{table}", payload)
        return int(saved["id"])

    def delete(self, table, row_id):
        if table in OWNED and self._lock_owned:
            return
        try:
            self._request("DELETE", f"/api/store/{table}/{int(row_id)}")
        except ApiError as exc:
            if getattr(exc, "status", None) == 404:
                return
            raise

    def offers_for(self, project_id):
        if self._lock_owned:
            return []
        return [self._record(item) for item in self._pages(f"/api/store/projects/{int(project_id)}/offers")]

    def offer_count(self, project_id):
        if self._lock_owned:
            return 0
        data = self._request("GET", f"/api/store/projects/{int(project_id)}/offer-count")
        return int(data["n"])

    def next_issue_no(self, project_id):
        data = self._request("GET", f"/api/store/projects/{int(project_id)}/next-issue")
        return int(data["n"])

    def delete_project_offers(self, project_id):
        if self._lock_owned:
            return
        if self.owner_id is not None and self.get("projects", project_id) is None:
            return
        for row in self.offers_for(project_id):
            self._remove_offer_file(row.get("file_path"))
        self._request("DELETE", f"/api/store/projects/{int(project_id)}/offers")

    def find_source(self, source_key):
        data = self._request("GET", "/api/store/invoices/source", query={"key": source_key})
        if not data or data.get("id") is None:
            return None
        return self._record(data)

    def talep_by_number(self, number):
        try:
            data = self._request("GET", "/api/store/talepler/number", query={"number": number})
        except ApiError as exc:
            if getattr(exc, "status", None) == 404:
                return None
            raise
        if not data:
            return None
        return self._record(data)

    def ignore(self, key):
        self._request("POST", "/api/store/ignored", {"key": key})

    def ignored(self, key):
        data = self._request("GET", "/api/store/ignored", query={"key": key})
        return bool(data["ignored"])

    def get_setting(self, key, default=""):
        data = self._request("GET", f"/api/settings/{key}")
        value = data.get("value")
        if value is None or value == "":
            return default
        return value

    def set_setting(self, key, value):
        self._request("PUT", f"/api/settings/{key}", {"value": value})

    def seen(self, key):
        data = self._request("GET", "/api/store/seen", query={"key": key})
        return bool(data["seen"])

    def mark_seen(self, key):
        self._request("POST", "/api/store/seen", {"key": key})

    def backup_to(self, path):
        url = self.base + "/api/store/backup"
        request = urllib.request.Request(url, headers=self._headers(), method="GET")
        try:
            with urllib.request.urlopen(request, timeout=60) as response:
                Path(path).write_bytes(response.read())
        except urllib.error.URLError as exc:
            raise ApiError("Yedek alınamadı.") from exc

    def _pages(self, path):
        items = []
        offset = 0
        while True:
            page = self._request("GET", path, query={"limit": 500, "offset": offset})
            batch = page["items"]
            items.extend(batch)
            if len(items) >= int(page["count"]) or not batch:
                break
            offset += len(batch)
        return items

    def _record(self, data):
        return Record(data)

    def _headers(self, auth=True):
        headers = {"Accept": "application/json"}
        if auth:
            headers["Authorization"] = f"Bearer {self.token}"
        if self.owner_id:
            headers["X-Owner-Id"] = str(self.owner_id)
        return headers

    def _request(self, method, path, body=None, query=None, auth=True):
        url = self.base + path
        if query:
            url += "?" + urllib.parse.urlencode({key: "" if value is None else value for key, value in query.items()})
        data = None
        headers = self._headers(auth)
        if body is not None:
            data = json.dumps(body, ensure_ascii=False).encode("utf-8")
            headers["Content-Type"] = "application/json"
        request = urllib.request.Request(url, data=data, headers=headers, method=method)
        try:
            with urllib.request.urlopen(request, timeout=30) as response:
                raw = response.read()
                if response.status == 204 or not raw:
                    return None
                return json.loads(raw.decode("utf-8"))
        except urllib.error.HTTPError as exc:
            detail = _detail(exc.read())
            if exc.code == 409:
                raise sqlite3.IntegrityError(detail)
            error = ApiError(detail or "İstek tamamlanamadı.")
            error.status = exc.code
            raise error from exc
        except urllib.error.URLError as exc:
            raise ApiError("Sunucuya ulaşılamadı.") from exc

    def _remove_offer_file(self, path):
        try:
            from is_merkezi import remove_offer_file
        except ImportError:
            return
        remove_offer_file(path)


def _detail(raw):
    text = raw.decode("utf-8", errors="replace")
    try:
        data = json.loads(text)
    except json.JSONDecodeError:
        return text.strip()
    detail = data.get("detail", text)
    if isinstance(detail, list):
        return "; ".join(str(item.get("msg", item)) if isinstance(item, dict) else str(item) for item in detail)
    return str(detail)
