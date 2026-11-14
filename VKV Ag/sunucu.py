# -*- coding: utf-8 -*-
"""VKV Ağ sunucusu. Yedek bilgisayarda bir kez açılır.

Her PC'deki istemci buraya bağlanır. Talepler, mesajlar ve dosyalar
yalnız bu veritabanında durur. İş Merkezi klasörüne yazılmaz.
"""
import json
import secrets
import sqlite3
import sys
import threading
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

from ortak import (
    PORT, ROLES, STATUSES, ONLINE_SECONDS, MAX_FILE,
    check_password, clean_username, hash_password, lan_ips, safe_filename, sunucu_dir,
)

def now_text():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def parse_time(value):
    try:
        return datetime.strptime(str(value or ""), "%Y-%m-%d %H:%M:%S")
    except ValueError:
        return None


class Hub:
    def __init__(self, root):
        self.root = Path(root)
        self.root.mkdir(parents=True, exist_ok=True)
        (self.root / "dosyalar").mkdir(parents=True, exist_ok=True)
        self.lock = threading.Lock()
        self.conn = sqlite3.connect(self.root / "ag.db", check_same_thread=False)
        self.conn.row_factory = sqlite3.Row
        self.conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS users (
              id INTEGER PRIMARY KEY,
              name TEXT NOT NULL,
              username TEXT NOT NULL UNIQUE,
              password TEXT NOT NULL,
              role TEXT NOT NULL,
              created_at TEXT
            );
            CREATE TABLE IF NOT EXISTS sessions (
              token TEXT PRIMARY KEY,
              user_id INTEGER NOT NULL,
              created_at TEXT
            );
            CREATE TABLE IF NOT EXISTS clients (
              istemci_id TEXT PRIMARY KEY,
              user_id INTEGER NOT NULL,
              cihaz TEXT,
              sistem TEXT,
              ip TEXT,
              last_seen TEXT
            );
            CREATE TABLE IF NOT EXISTS requests (
              id INTEGER PRIMARY KEY,
              number TEXT NOT NULL UNIQUE,
              title TEXT NOT NULL,
              note TEXT,
              author_id INTEGER NOT NULL,
              assignee_id INTEGER NOT NULL,
              status TEXT NOT NULL,
              created_at TEXT
            );
            CREATE TABLE IF NOT EXISTS files (
              id INTEGER PRIMARY KEY,
              name TEXT NOT NULL,
              stored TEXT NOT NULL,
              size INTEGER NOT NULL,
              note TEXT,
              owner_id INTEGER NOT NULL,
              created_at TEXT
            );
            CREATE TABLE IF NOT EXISTS messages (
              id INTEGER PRIMARY KEY,
              user_id INTEGER NOT NULL,
              text TEXT NOT NULL,
              created_at TEXT
            );
            """
        )
        self.conn.commit()
        self._ensure_key()

    def close(self):
        self.conn.close()

    def _ensure_key(self):
        path = self.root / "kurulum.anahtar"
        if self.user_count():
            try:
                path.unlink()
            except OSError:
                pass
            self.setup_key = ""
            return
        if path.is_file():
            self.setup_key = path.read_text(encoding="utf-8").strip()
        else:
            self.setup_key = secrets.token_hex(4)
            path.write_text(self.setup_key + "\n", encoding="utf-8")
            try:
                path.chmod(0o600)
            except OSError:
                pass

    def user_count(self):
        return self.conn.execute("SELECT COUNT(*) AS n FROM users").fetchone()["n"]

    def needs_setup(self):
        return self.user_count() == 0

    def setup(self, key, name, username, password):
        with self.lock:
            if not self.needs_setup():
                raise ValueError("Kurulum tamamlanmış.")
            given = str(key or "").strip()
            expected = self.setup_key or ""
            if len(given) != len(expected) or not secrets.compare_digest(given, expected):
                raise ValueError("Kurulum anahtarı uyuşmuyor.")
            user_id = self._add_user(name, username, password, "Yönetici")
            self._ensure_key()
            return self._public_user(self._user(user_id))

    def _add_user(self, name, username, password, role):
        name = " ".join(str(name or "").split())
        username = clean_username(username)
        if len(name) < 2 or len(username) < 3 or len(str(password or "")) < 6:
            raise ValueError("Ad, en az 3 karakterlik kullanıcı adı ve 6 karakterlik parola gerekli.")
        if role not in ROLES:
            raise ValueError("Rol seç.")
        try:
            cursor = self.conn.execute(
                "INSERT INTO users (name, username, password, role, created_at) VALUES (?,?,?,?,?)",
                (name[:80], username, hash_password(password), role, now_text()),
            )
        except sqlite3.IntegrityError as exc:
            raise ValueError("Bu kullanıcı adı kayıtlı.") from exc
        self.conn.commit()
        return cursor.lastrowid

    def add_user(self, actor, name, username, password, role):
        if actor["rol"] != "Yönetici":
            raise ValueError("Kullanıcıyı yönetici ekler.")
        with self.lock:
            return self._public_user(self._user(self._add_user(name, username, password, role)))

    def login(self, username, password):
        username = clean_username(username)
        row = self.conn.execute("SELECT * FROM users WHERE username=?", (username,)).fetchone()
        if row is None or not check_password(password, row["password"]):
            raise ValueError("Kullanıcı adı veya parola hatalı.")
        token = secrets.token_urlsafe(24)
        with self.lock:
            self.conn.execute(
                "INSERT INTO sessions (token, user_id, created_at) VALUES (?,?,?)",
                (token, row["id"], now_text()),
            )
            self.conn.commit()
        user = self._public_user(row)
        user["token"] = token
        return user

    def logout(self, token):
        with self.lock:
            self.conn.execute("DELETE FROM sessions WHERE token=?", (token,))
            self.conn.commit()

    def user_by_token(self, token):
        if not token:
            return None
        row = self.conn.execute(
            """
            SELECT u.* FROM sessions s
            JOIN users u ON u.id = s.user_id
            WHERE s.token=?
            """,
            (token,),
        ).fetchone()
        return self._public_user(row) if row else None

    def _user(self, user_id):
        return self.conn.execute("SELECT * FROM users WHERE id=?", (int(user_id),)).fetchone()

    def _public_user(self, row):
        if row is None:
            return None
        return {"id": row["id"], "ad": row["name"], "kullanici": row["username"], "rol": row["role"]}

    def users(self):
        rows = self.conn.execute("SELECT * FROM users ORDER BY name").fetchall()
        return [self._public_user(row) for row in rows]

    def beat(self, user, istemci_id, cihaz, sistem, ip):
        istemci_id = "".join(ch for ch in str(istemci_id or "") if ch.isalnum() or ch in "-")[:40]
        if len(istemci_id) < 8:
            raise ValueError("İstemci kimliği yok.")
        with self.lock:
            self.conn.execute(
                """
                INSERT INTO clients (istemci_id, user_id, cihaz, sistem, ip, last_seen)
                VALUES (?,?,?,?,?,?)
                ON CONFLICT(istemci_id) DO UPDATE SET
                  user_id=excluded.user_id, cihaz=excluded.cihaz, sistem=excluded.sistem,
                  ip=excluded.ip, last_seen=excluded.last_seen
                """,
                (istemci_id, user["id"], str(cihaz or "")[:80], str(sistem or "")[:80], str(ip or "")[:64], now_text()),
            )
            self.conn.commit()
        return {"ok": True}

    def computers(self):
        rows = self.conn.execute(
            """
            SELECT c.*, u.name AS user_name, u.username, u.role
            FROM clients c JOIN users u ON u.id = c.user_id
            ORDER BY c.last_seen DESC
            """
        ).fetchall()
        moment = datetime.now()
        listed = []
        for row in rows:
            seen = parse_time(row["last_seen"])
            online = seen is not None and (moment - seen).total_seconds() <= ONLINE_SECONDS
            listed.append({
                "istemci": row["istemci_id"],
                "kullanici_id": row["user_id"],
                "ad": row["user_name"],
                "kullanici": row["username"],
                "rol": row["role"],
                "cihaz": row["cihaz"],
                "sistem": row["sistem"],
                "ip": row["ip"],
                "gorulme": row["last_seen"],
                "cevrimici": online,
            })
        return listed

    def create_request(self, actor, title, note, assignee_id):
        title = " ".join(str(title or "").split())
        note = " ".join(str(note or "").split())
        if len(title) < 2:
            raise ValueError("Başlık gerekli.")
        try:
            assignee_id = int(assignee_id)
        except (TypeError, ValueError) as exc:
            raise ValueError("Talep verilecek kişiyi seç.") from exc
        assignee = self._user(assignee_id)
        if assignee is None:
            raise ValueError("Talep verilecek kişiyi seç.")
        with self.lock:
            count = self.conn.execute("SELECT COUNT(*) AS n FROM requests").fetchone()["n"]
            number = f"AG-{int(count) + 1}"
            while self.conn.execute("SELECT id FROM requests WHERE number=?", (number,)).fetchone():
                count += 1
                number = f"AG-{int(count) + 1}"
            cursor = self.conn.execute(
                """
                INSERT INTO requests (number, title, note, author_id, assignee_id, status, created_at)
                VALUES (?,?,?,?,?,?,?)
                """,
                (number, title[:180], note[:2000], actor["id"], assignee["id"], "Açık", now_text()),
            )
            self.conn.commit()
            request_id = cursor.lastrowid
        return self._request(request_id)

    def set_status(self, actor, request_id, status):
        if status not in STATUSES:
            raise ValueError("Durum yok.")
        try:
            request_id = int(request_id)
        except (TypeError, ValueError) as exc:
            raise ValueError("Talep yok.") from exc
        row = self.conn.execute("SELECT * FROM requests WHERE id=?", (request_id,)).fetchone()
        if row is None:
            raise ValueError("Talep yok.")
        if actor["rol"] != "Yönetici" and actor["id"] not in {row["author_id"], row["assignee_id"]}:
            raise ValueError("Bu talebin durumunu açan veya alan kişi değiştirir.")
        with self.lock:
            self.conn.execute("UPDATE requests SET status=? WHERE id=?", (status, request_id))
            self.conn.commit()
        return self._request(request_id)

    def requests_for(self, actor):
        if actor["rol"] == "Yönetici":
            rows = self.conn.execute(self._request_sql() + " ORDER BY r.id DESC").fetchall()
        else:
            rows = self.conn.execute(
                self._request_sql() + " WHERE r.author_id=? OR r.assignee_id=? ORDER BY r.id DESC",
                (actor["id"], actor["id"]),
            ).fetchall()
        return [self._request_dict(row) for row in rows]

    def distribution(self):
        people = []
        for user in self.users():
            open_count = self.conn.execute(
                "SELECT COUNT(*) AS n FROM requests WHERE assignee_id=? AND status!='Tamamlandı'",
                (user["id"],),
            ).fetchone()["n"]
            people.append({**user, "acik": open_count})
        return people

    def _request(self, request_id):
        row = self.conn.execute(self._request_sql() + " WHERE r.id=?", (int(request_id),)).fetchone()
        return self._request_dict(row)

    def _request_sql(self):
        return """
            SELECT r.*, a.name AS author_name, b.name AS assignee_name
            FROM requests r
            JOIN users a ON a.id = r.author_id
            JOIN users b ON b.id = r.assignee_id
        """

    def _request_dict(self, row):
        return {
            "id": row["id"],
            "no": row["number"],
            "baslik": row["title"],
            "not": row["note"] or "",
            "acan": row["author_name"],
            "acan_id": row["author_id"],
            "alan": row["assignee_name"],
            "alan_id": row["assignee_id"],
            "durum": row["status"],
            "zaman": row["created_at"],
        }

    def save_file(self, actor, name, note, payload):
        if not payload:
            raise ValueError("Dosya boş.")
        if len(payload) > MAX_FILE:
            raise ValueError("Dosya 20 MB sınırını aşıyor.")
        clean = safe_filename(name)
        with self.lock:
            cursor = self.conn.execute(
                "INSERT INTO files (name, stored, size, note, owner_id, created_at) VALUES (?,?,?,?,?,?)",
                (clean, "", len(payload), " ".join(str(note or "").split())[:200], actor["id"], now_text()),
            )
            file_id = cursor.lastrowid
            stored = f"{file_id}-{clean}"
            path = self.root / "dosyalar" / stored
            path.write_bytes(payload)
            self.conn.execute("UPDATE files SET stored=? WHERE id=?", (stored, file_id))
            self.conn.commit()
        return self._file_dict(file_id)

    def files(self):
        rows = self.conn.execute(
            """
            SELECT f.*, u.name AS owner_name
            FROM files f JOIN users u ON u.id = f.owner_id
            ORDER BY f.id DESC
            """
        ).fetchall()
        return [self._file_row(row) for row in rows]

    def file_bytes(self, file_id):
        try:
            file_id = int(file_id)
        except (TypeError, ValueError) as exc:
            raise ValueError("Dosya yok.") from exc
        row = self.conn.execute("SELECT * FROM files WHERE id=?", (file_id,)).fetchone()
        if row is None:
            raise ValueError("Dosya yok.")
        path = (self.root / "dosyalar" / row["stored"]).resolve()
        folder = (self.root / "dosyalar").resolve()
        if folder not in path.parents or not path.is_file():
            raise ValueError("Dosya bulunamadı.")
        return row["name"], path.read_bytes()

    def _file_dict(self, file_id):
        row = self.conn.execute(
            """
            SELECT f.*, u.name AS owner_name FROM files f
            JOIN users u ON u.id = f.owner_id WHERE f.id=?
            """,
            (int(file_id),),
        ).fetchone()
        return self._file_row(row)

    def _file_row(self, row):
        return {
            "id": row["id"],
            "ad": row["name"],
            "boy": row["size"],
            "not": row["note"] or "",
            "sahip": row["owner_name"],
            "sahip_id": row["owner_id"],
            "zaman": row["created_at"],
        }

    def add_message(self, actor, text):
        text = " ".join(str(text or "").split())
        if not text:
            raise ValueError("Mesaj yaz.")
        with self.lock:
            cursor = self.conn.execute(
                "INSERT INTO messages (user_id, text, created_at) VALUES (?,?,?)",
                (actor["id"], text[:500], now_text()),
            )
            self.conn.commit()
            message_id = cursor.lastrowid
        return self._message(message_id)

    def messages(self):
        rows = self.conn.execute(
            """
            SELECT m.*, u.name AS user_name FROM messages m
            JOIN users u ON u.id = m.user_id
            ORDER BY m.id DESC LIMIT 80
            """
        ).fetchall()
        return [self._message_row(row) for row in reversed(rows)]

    def _message(self, message_id):
        row = self.conn.execute(
            """
            SELECT m.*, u.name AS user_name FROM messages m
            JOIN users u ON u.id = m.user_id WHERE m.id=?
            """,
            (int(message_id),),
        ).fetchone()
        return self._message_row(row)

    def _message_row(self, row):
        return {"id": row["id"], "ad": row["user_name"], "metin": row["text"], "zaman": row["created_at"]}


class Handler(BaseHTTPRequestHandler):
    protocol_version = "HTTP/1.1"

    def log_message(self, _format, *_args):
        return

    def do_GET(self):
        self._handle()

    def do_POST(self):
        self._handle()

    def _handle(self):
        path = urllib.parse.urlsplit(self.path).path
        try:
            if path == "/v1/durum" and self.command == "GET":
                self._json(200, {"ok": True, "kurulum": self.server.hub.needs_setup(), "ad": "VKV Ağ"})
                return
            if path == "/v1/kurulum" and self.command == "POST":
                body = self._json_body()
                user = self.server.hub.setup(body.get("anahtar"), body.get("ad"), body.get("kullanici"), body.get("parola"))
                self._json(200, {"ok": True, "kullanici": user})
                return
            if path == "/v1/giris" and self.command == "POST":
                body = self._json_body()
                self._json(200, {"ok": True, "kullanici": self.server.hub.login(body.get("kullanici"), body.get("parola"))})
                return
            user = self._actor()
            if user is None:
                self._json(401, {"ok": False, "error": "Oturum kapandı. Yeniden gir."})
                return
            if path == "/v1/cikis" and self.command == "POST":
                self.server.hub.logout(self._token())
                self._json(200, {"ok": True})
                return
            if path == "/v1/nabiz" and self.command == "POST":
                body = self._json_body()
                peer = self.client_address[0]
                self.server.hub.beat(user, self.headers.get("X-Istemci"), body.get("cihaz"), body.get("sistem"), peer or body.get("ip"))
                self._json(200, {"ok": True})
                return
            if path == "/v1/kullanicilar" and self.command == "GET":
                self._json(200, {"ok": True, "kullanicilar": self.server.hub.users()})
                return
            if path == "/v1/kullanici" and self.command == "POST":
                body = self._json_body()
                created = self.server.hub.add_user(user, body.get("ad"), body.get("kullanici"), body.get("parola"), body.get("rol") or "Kullanıcı")
                self._json(200, {"ok": True, "kullanici": created})
                return
            if path == "/v1/bilgisayarlar" and self.command == "GET":
                self._json(200, {"ok": True, "bilgisayarlar": self.server.hub.computers()})
                return
            if path == "/v1/dagilim" and self.command == "GET":
                self._json(200, {"ok": True, "dagilim": self.server.hub.distribution()})
                return
            if path == "/v1/talepler" and self.command == "GET":
                self._json(200, {"ok": True, "talepler": self.server.hub.requests_for(user)})
                return
            if path == "/v1/talep" and self.command == "POST":
                body = self._json_body()
                created = self.server.hub.create_request(user, body.get("baslik"), body.get("not"), body.get("kime"))
                self._json(200, {"ok": True, "talep": created})
                return
            if path == "/v1/talep/durum" and self.command == "POST":
                body = self._json_body()
                updated = self.server.hub.set_status(user, body.get("id"), body.get("durum"))
                self._json(200, {"ok": True, "talep": updated})
                return
            if path == "/v1/dosyalar" and self.command == "GET":
                self._json(200, {"ok": True, "dosyalar": self.server.hub.files()})
                return
            if path == "/v1/dosya" and self.command == "POST":
                length = self._length(MAX_FILE)
                payload = self.rfile.read(length) if length else b""
                name = urllib.parse.unquote(self.headers.get("X-Dosya-Ad", ""))
                note = urllib.parse.unquote(self.headers.get("X-Dosya-Not", ""))
                saved = self.server.hub.save_file(user, name, note, payload)
                self._json(200, {"ok": True, "dosya": saved})
                return
            if path == "/v1/dosya" and self.command == "GET":
                query = urllib.parse.parse_qs(urllib.parse.urlsplit(self.path).query)
                name, payload = self.server.hub.file_bytes((query.get("id") or ["0"])[0])
                quoted = urllib.parse.quote(name)
                self.send_response(200)
                self.send_header("Content-Type", "application/octet-stream")
                self.send_header("Content-Length", str(len(payload)))
                self.send_header("Content-Disposition", f"attachment; filename*=UTF-8''{quoted}")
                self.end_headers()
                self.wfile.write(payload)
                return
            if path == "/v1/mesajlar" and self.command == "GET":
                self._json(200, {"ok": True, "mesajlar": self.server.hub.messages()})
                return
            if path == "/v1/mesaj" and self.command == "POST":
                body = self._json_body()
                self._json(200, {"ok": True, "mesaj": self.server.hub.add_message(user, body.get("metin"))})
                return
            self._json(404, {"ok": False, "error": "Yok."})
        except ValueError as exc:
            self._json(400, {"ok": False, "error": str(exc)})
        except Exception:
            self._json(500, {"ok": False, "error": "Sunucu isteği tamamlayamadı."})

    def _actor(self):
        return self.server.hub.user_by_token(self._token())

    def _token(self):
        header = self.headers.get("Authorization", "")
        return header[7:].strip() if header.startswith("Bearer ") else ""

    def _length(self, limit):
        try:
            length = int(self.headers.get("Content-Length", "0"))
        except ValueError as exc:
            raise ValueError("İstek boyu okunamadı.") from exc
        if length < 0 or length > limit:
            raise ValueError("İstek boyu geçersiz.")
        return length

    def _json_body(self):
        length = self._length(64 * 1024)
        if length == 0:
            return {}
        try:
            data = json.loads(self.rfile.read(length).decode("utf-8"))
        except (UnicodeError, json.JSONDecodeError) as exc:
            raise ValueError("İstek okunamadı.") from exc
        if not isinstance(data, dict):
            raise ValueError("İstek okunamadı.")
        return data

    def _json(self, status, payload):
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)


class Server(ThreadingHTTPServer):
    allow_reuse_address = True

    def __init__(self, address, hub):
        self.hub = hub
        super().__init__(address, Handler)


def serve(root, host, port):
    hub = Hub(root)
    httpd = Server((host, port), hub)
    ips = lan_ips()
    print("VKV Ağ sunucusu açık.")
    print(f"Veri: {root}")
    if ips:
        print("Diğer bilgisayarlar bu adresi yazar:")
        for ip in ips:
            print(f"  {ip}:{port}")
    print(f"Aynı bilgisayardan: 127.0.0.1:{port}")
    if hub.needs_setup():
        print(f"İlk yönetici anahtarı: {hub.setup_key}")
    else:
        print("Yönetici kurulu. İstemciler kullanıcı adı ve parola ile girer.")
    print("Kapatmak için Ctrl+C.")
    sys.stdout.flush()
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nSunucu kapandı.")
    finally:
        httpd.server_close()
        hub.close()


def _self_check():
    import tempfile
    import urllib.request
    folder = tempfile.TemporaryDirectory()
    hub = Hub(Path(folder.name))
    httpd = Server(("127.0.0.1", 0), hub)
    thread = threading.Thread(target=httpd.serve_forever, daemon=True)
    thread.start()
    port = httpd.server_address[1]
    base = f"http://127.0.0.1:{port}"

    def call(path, payload=None, token="", raw=None, headers=None):
        data = raw if raw is not None else (None if payload is None else json.dumps(payload).encode("utf-8"))
        request = urllib.request.Request(base + path, data=data, method="GET" if data is None else "POST")
        if payload is not None and raw is None:
            request.add_header("Content-Type", "application/json")
        if token:
            request.add_header("Authorization", f"Bearer {token}")
        for key, value in (headers or {}).items():
            request.add_header(key, value)
        try:
            with urllib.request.urlopen(request, timeout=5) as response:
                body = response.read()
                kind = response.headers.get("Content-Type", "")
                return response.status, json.loads(body.decode("utf-8")) if "json" in kind else body
        except urllib.error.HTTPError as exc:
            body = exc.read()
            try:
                return exc.code, json.loads(body.decode("utf-8"))
            except json.JSONDecodeError:
                return exc.code, body

    try:
        status, body = call("/v1/durum")
        assert status == 200 and body["kurulum"]
        status, body = call("/v1/kurulum", {"anahtar": "yanlis", "ad": "Furkan", "kullanici": "furkan", "parola": "parola1"})
        assert status == 400
        status, body = call("/v1/kurulum", {"anahtar": hub.setup_key, "ad": "Furkan Hamarat", "kullanici": "furkan", "parola": "parola1"})
        assert status == 200 and body["kullanici"]["rol"] == "Yönetici"
        assert not hub.needs_setup()
        status, admin = call("/v1/giris", {"kullanici": "furkan", "parola": "parola1"})
        token = admin["kullanici"]["token"]
        status, created = call("/v1/kullanici", {"ad": "Ali Veli", "kullanici": "ali", "parola": "parola2", "rol": "Kullanıcı"}, token)
        assert created["kullanici"]["kullanici"] == "ali"
        status, other = call("/v1/giris", {"kullanici": "ali", "parola": "parola2"})
        other_token = other["kullanici"]["token"]
        status, denied = call("/v1/kullanici", {"ad": "Veli", "kullanici": "veli", "parola": "parola3", "rol": "Kullanıcı"}, other_token)
        assert status == 400 and denied["ok"] is False
        status, beat = call("/v1/nabiz", {"cihaz": "Ofis", "sistem": "macOS", "ip": "10.0.0.5"}, token, headers={"X-Istemci": "istemci-1"})
        assert beat["ok"]
        status, machines = call("/v1/bilgisayarlar", token=token)
        assert machines["bilgisayarlar"][0]["cevrimici"]
        assert machines["bilgisayarlar"][0]["ip"]
        ali_id = created["kullanici"]["id"]
        status, talep = call("/v1/talep", {"baslik": "Fatura kontrol", "not": "Mart klasörü", "kime": ali_id}, token)
        assert talep["talep"]["no"] == "AG-1"
        assert talep["talep"]["alan"] == "Ali Veli"
        status, seen = call("/v1/talepler", token=other_token)
        assert seen["talepler"][0]["no"] == "AG-1"
        status, moved = call("/v1/talep/durum", {"id": talep["talep"]["id"], "durum": "Üzerinde"}, other_token)
        assert moved["talep"]["durum"] == "Üzerinde"
        status, spread = call("/v1/dagilim", token=token)
        ali = next(item for item in spread["dagilim"] if item["kullanici"] == "ali")
        assert ali["acik"] == 1
        payload = b"%PDF-1.4 deneme"
        status, saved = call(
            "/v1/dosya", token=token, raw=payload,
            headers={"X-Dosya-Ad": urllib.parse.quote("Mart fatura.pdf"), "X-Dosya-Not": urllib.parse.quote("ortak")},
        )
        assert saved["dosya"]["ad"] == "Mart fatura.pdf"
        status, downloaded = call(f"/v1/dosya?id={saved['dosya']['id']}", token=token)
        assert downloaded == payload
        status, message = call("/v1/mesaj", {"metin": "Dosya duruyor"}, other_token)
        status, board = call("/v1/mesajlar", token=token)
        assert board["mesajlar"][-1]["metin"] == "Dosya duruyor"
        call("/v1/cikis", {}, token)
        status, closed = call("/v1/talepler", token=token)
        assert status == 401
        assert message["ok"] and status == 401
    finally:
        httpd.shutdown()
        httpd.server_close()
        hub.close()
        folder.cleanup()
    print("ok")


def main():
    if "--check" in sys.argv:
        _self_check()
        return
    root = sunucu_dir()
    serve(root, "0.0.0.0", PORT)


if __name__ == "__main__":
    main()
