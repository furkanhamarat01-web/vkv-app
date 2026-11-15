"""VKV Talep arayüzünün veri katmanı. Kayıtlar İş Merkezi API'sinde durur."""

import base64
import json
import os
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path


class ApiError(RuntimeError):
    def __init__(self, message, status=None):
        super().__init__(message)
        self.status = status


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


def _cache_dir():
    root = Path.home() / "Library" / "Application Support" / "VKVTalep" / "Dosyalar"
    root.mkdir(parents=True, exist_ok=True)
    return root


class Store:
    def __init__(self):
        self.base = _base()
        self.token = _token()
        if not self.token:
            raise ApiError("API anahtarı yok. Uygulamayı İş Merkezi simgesinden bir kez açın.")
        try:
            self._request("GET", "/health", auth=False)
        except ApiError as exc:
            raise ApiError(f"Sunucuya bağlanılamadı ({self.base}).\n{exc}") from exc

    def close(self):
        return None

    def user_count(self):
        return int(self._request("GET", "/api/talep/users/count")["n"])

    def setup(self, name, username, password, email=""):
        return self._request("POST", "/api/talep/setup", {
            "name": name, "username": username, "password": password, "email": email,
        })

    def session(self, username, password):
        return self._request("POST", "/api/talep/session", {"username": username, "password": password})["user"]

    def forgot(self, username):
        self._request("POST", "/api/talep/forgot", {"username": username})

    def account(self, user_id):
        return self._request("GET", f"/api/talep/users/{int(user_id)}/account")

    def resets(self):
        return self._request("GET", "/api/talep/resets")

    def grant_reset(self, reset_id, password):
        self._request("POST", f"/api/talep/resets/{int(reset_id)}/password", {"password": password, "must_change": True})

    def team_members(self, name):
        return self._request("GET", "/api/talep/teams/members", query={"name": name})

    def add_user(self, name, email, password, secret, is_admin=False, username="", must_change=None, role=None):
        body = {
            "name": name,
            "email": email,
            "password": password,
            "secret": secret,
            "is_admin": bool(is_admin),
            "username": username or "",
            "role": role,
        }
        if must_change is not None:
            body["must_change"] = bool(must_change)
        return int(self._request("POST", "/api/talep/users", body)["id"])

    def set_password(self, user_id, password, must_change=False):
        self._request("POST", f"/api/talep/users/{int(user_id)}/password", {
            "password": password,
            "must_change": bool(must_change),
        })

    def update_user(self, user_id, name, email, username, team_id=None, role=None):
        self._request("PATCH", f"/api/talep/users/{int(user_id)}", {
            "name": name,
            "email": email,
            "username": username,
            "team_id": int(team_id) if team_id else None,
            "role": role,
        })

    def add_login_log(self, email, event, user=None, device=None):
        self._request("POST", "/api/talep/login-logs", {
            "email": email,
            "event": event,
            "user": None if user is None else {"id": user["id"], "name": user.get("name") or ""},
            "device": device or {},
        })

    def login_logs(self, limit=400):
        return self._request("GET", "/api/talep/login-logs", query={"limit": int(limit)})

    def user_by_email(self, email):
        return self._request("GET", "/api/talep/users/by-email", query={"email": email})

    def users(self):
        return self._request("GET", "/api/talep/users")

    def add_team(self, name, code, topic=""):
        return int(self._request("POST", "/api/talep/teams", {"name": name, "code": code, "topic": topic})["id"])

    def update_team(self, team_id, name, topic, code):
        self._request("PATCH", f"/api/talep/teams/{int(team_id)}", {"name": name, "topic": topic, "code": code})

    def teams(self):
        return self._request("GET", "/api/talep/teams")

    def team_by_code(self, code):
        return self._request("GET", "/api/talep/teams/by-code", query={"code": code})

    def add_request_code(self, code, team_id, work_code, title="", scope="İç", all_teams=False):
        return int(self._request("POST", "/api/talep/request-codes", {
            "code": code,
            "team_id": int(team_id),
            "work_code": work_code,
            "title": title,
            "scope": scope,
            "all_teams": bool(all_teams),
        })["id"])

    def update_request_code(self, code_id, code, team_id, work_code, title="", scope="İç", all_teams=False):
        self._request("PATCH", f"/api/talep/request-codes/{int(code_id)}", {
            "code": code,
            "team_id": int(team_id),
            "work_code": work_code,
            "title": title,
            "scope": scope,
            "all_teams": bool(all_teams),
        })

    def request_codes(self):
        return self._request("GET", "/api/talep/request-codes")

    def add_member(self, user_id, team_id):
        self._request("POST", "/api/talep/members", {"user_id": int(user_id), "team_id": int(team_id)})

    def team_ids(self, user_id):
        return {int(item) for item in self._request("GET", f"/api/talep/users/{int(user_id)}/team-ids")["ids"]}

    def user_teams(self, user_id):
        return self._request("GET", f"/api/talep/users/{int(user_id)}/teams")

    def member_count(self, team_id):
        return int(self._request("GET", f"/api/talep/teams/{int(team_id)}/member-count")["n"])

    def login(self, email, password, code):
        return self._request("POST", "/api/talep/login", {"email": email, "password": password, "code": code})["user"]

    def create_request(self, author_id, team_id, target_team_id, work_code, title, description, priority, due_date):
        return self._request("POST", "/api/talep/requests", {
            "author_id": int(author_id),
            "team_id": int(team_id),
            "target_team_id": int(target_team_id),
            "work_code": work_code,
            "title": title,
            "description": description,
            "priority": priority,
            "due_date": due_date,
        })

    def attach(self, request_id, src):
        source = Path(src)
        payload = source.read_bytes()
        row = self._request("POST", f"/api/talep/requests/{int(request_id)}/file", {
            "name": source.name,
            "data": base64.b64encode(payload).decode("ascii"),
        })
        return self._with_file(row)

    def pending_vkv(self):
        return [self._with_file(row) for row in self._request("GET", "/api/talep/requests/pending-vkv")]

    def mark_vkv(self, request_id, task_id, issue_key):
        self._request("POST", f"/api/talep/requests/{int(request_id)}/vkv", {
            "task_id": int(task_id) if task_id else None,
            "issue_key": issue_key or "",
        })

    def set_status(self, request_id, status, note=""):
        self._request("POST", f"/api/talep/requests/{int(request_id)}/status", {"status": status, "note": note})

    def request(self, request_id):
        return self._with_file(self._request("GET", f"/api/talep/requests/{int(request_id)}"))

    def visible_requests(self, user):
        rows = self._request("GET", "/api/talep/requests", query={"user_id": int(user["id"])})
        return [self._with_file(row) for row in rows]

    def _with_file(self, row):
        if not row or not row.get("file_path"):
            return row
        name = Path(str(row["file_path"])).name
        dest = _cache_dir() / f"{int(row['id'])}-{name}"
        if not dest.is_file():
            url = self.base + f"/api/talep/requests/{int(row['id'])}/file"
            request = urllib.request.Request(url, headers=self._headers(), method="GET")
            try:
                with urllib.request.urlopen(request, timeout=60) as response:
                    dest.write_bytes(response.read())
            except urllib.error.URLError:
                return row
        copied = dict(row)
        copied["file_path"] = str(dest)
        return copied

    def _headers(self, auth=True):
        headers = {"Accept": "application/json"}
        if auth:
            headers["Authorization"] = f"Bearer {self.token}"
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
                kind = response.headers.get("Content-Type", "")
                if response.status == 204 or not raw:
                    return None
                if "json" not in kind:
                    return raw
                return json.loads(raw.decode("utf-8"))
        except urllib.error.HTTPError as exc:
            detail = _detail(exc.read())
            if exc.code == 400:
                raise ValueError(detail or "İstek geçersiz.") from exc
            error = ApiError(detail or "İstek tamamlanamadı.")
            error.status = exc.code
            raise error from exc
        except urllib.error.URLError as exc:
            raise ApiError("Sunucuya ulaşılamadı.") from exc


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
