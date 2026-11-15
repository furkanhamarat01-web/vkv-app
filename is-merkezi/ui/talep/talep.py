#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""VKV Talep — VKV İş Takip'ten ayrı ekip talep programı.

Kayıtlar İş Merkezi API'sinde durur. VKV aynı bilgisayarda açıksa talep, iş takibine de düşer.
"""
import base64
import hashlib
import hmac
import json
import os
import platform
import secrets
import shutil
import socket
import subprocess
import sys
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime
from pathlib import Path
import tkinter as tk

sys.path.insert(0, str(Path(__file__).resolve().parent))
from talep_client import ApiError, Store
from tkinter import filedialog, messagebox, ttk

APP_NAME = "Talep Yönetimi"
FONT = ".AppleSystemUIFont" if sys.platform == "darwin" else "Segoe UI"
NAV = "#1D2438"
NAV_TEXT = "#AEB6CC"
ACCENT = "#3250E6"
ACCENT_HOVER = "#2843C4"
CANVAS = "#F3F4F7"
INK = "#14182B"
MUTED = "#6B7287"
CARD = "#FFFFFF"
LINE = "#E3E6EE"
SOFT = "#EAEEFD"
HOVER = "#EEF0F4"
OK = "#1F8A4C"
DANGER = "#D1343B"
PRIORITIES = ("Düşük", "Normal", "Yüksek", "Acil")
STATUSES = ("Açık", "Üzerinde", "Beklemede", "Tamamlandı", "Reddedildi")
NOTE_STATUSES = ("Beklemede", "Tamamlandı", "Reddedildi")
ADMIN_EMAIL = "furkan.hamarat01@gmail.com"
ADMIN_ROLE = "Yönetici"
ROLES = ("Kullanıcı", "Sorumlu", "İzleyici")
LOG_EVENTS = ("Giriş", "Çıkış", "Başarısız")
FILE_EXT = {".pdf", ".png", ".jpg", ".jpeg"}
MAX_FILE = 20 * 1024 * 1024


def app_dir():
    if sys.platform == "darwin":
        root = Path.home() / "Library" / "Application Support" / "VKVTalep"
    elif sys.platform == "win32":
        root = Path(os.environ.get("APPDATA", str(Path.home()))) / "VKVTalep"
    else:
        root = Path.home() / ".vkv_talep"
    root.mkdir(parents=True, exist_ok=True)
    return root


def session_path():
    if sys.platform == "darwin":
        return Path.home() / "Library" / "Application Support" / "IsMerkezi" / "talep.oturum"
    if sys.platform == "win32":
        return Path(os.environ.get("APPDATA", str(Path.home()))) / "IsMerkezi" / "talep.oturum"
    return Path.home() / ".is_merkezi" / "talep.oturum"


def write_session(user):
    path = session_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(f"{user['id']}\n{user['email']}\n", encoding="utf-8")
    try:
        os.chmod(path, 0o600)
    except OSError:
        pass


def clear_session():
    try:
        session_path().unlink()
    except OSError:
        pass


def kopru_path():
    if sys.platform == "darwin":
        return Path.home() / "Library" / "Application Support" / "IsMerkezi" / "talep.kopru"
    if sys.platform == "win32":
        return Path(os.environ.get("APPDATA", str(Path.home()))) / "IsMerkezi" / "talep.kopru"
    return Path.home() / ".is_merkezi" / "talep.kopru"


def hash_password(password, salt=None):
    salt = salt or secrets.token_bytes(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, 120_000)
    return salt.hex() + ":" + digest.hex()


def check_password(password, stored):
    try:
        salt_hex, digest_hex = str(stored).split(":", 1)
        salt = bytes.fromhex(salt_hex)
        digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, 120_000)
    except (ValueError, TypeError):
        return False
    return secrets.compare_digest(digest.hex(), digest_hex)


def new_totp_secret():
    return base64.b32encode(secrets.token_bytes(20)).decode("ascii").rstrip("=")


def totp_codes(secret, moment=None, window=1):
    key = base64.b32decode(secret.upper() + "=" * ((8 - len(secret) % 8) % 8), casefold=True)
    counter = int((time.time() if moment is None else moment) // 30)
    codes = []
    for offset in range(-window, window + 1):
        digest = hmac.new(key, (counter + offset).to_bytes(8, "big"), hashlib.sha1).digest()
        start = digest[-1] & 0x0F
        number = int.from_bytes(digest[start:start + 4], "big") & 0x7FFFFFFF
        codes.append(f"{number % 1_000_000:06d}")
    return codes


def totp_ok(secret, code):
    typed = "".join(ch for ch in str(code or "") if ch.isdigit())
    if len(typed) != 6:
        return False
    return typed in totp_codes(secret)


def otpauth_uri(email, secret):
    label = urllib.parse.quote(f"VKV Talep:{email}")
    return f"otpauth://totp/{label}?secret={secret}&issuer=VKV%20Talep&digits=6&period=30"


def clean_code(value):
    return "".join(ch for ch in str(value or "").upper() if ch.isalnum())[:8]


def clean_request_code(value):
    return "".join(ch for ch in str(value or "").upper() if ch.isalnum() or ch == "-")[:16]


def clean_email(value):
    return " ".join(str(value or "").split()).casefold()


def clean_username(value):
    return "".join(ch for ch in str(value or "").casefold() if ch.isalnum() or ch in "._-")[:32]


def device_snapshot():
    host = platform.node() or ""
    if sys.platform == "darwin":
        try:
            named = subprocess.check_output(["scutil", "--get", "ComputerName"], text=True, timeout=2).strip()
            if named:
                host = named
        except (OSError, subprocess.SubprocessError):
            pass
    elif sys.platform == "win32":
        host = os.environ.get("COMPUTERNAME") or host
    system = "macOS" if platform.system() == "Darwin" else platform.system()
    os_text = " ".join(part for part in (system, platform.release(), platform.machine()) if part)
    ip = ""
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as sock:
            sock.settimeout(0.3)
            sock.connect(("1.1.1.1", 80))
            ip = sock.getsockname()[0]
    except OSError:
        ip = ""
    return {
        "device": (host or "Bu bilgisayar")[:80],
        "os": os_text[:80],
        "ip": ip[:64],
    }


def fold_ascii(value):
    return str(value or "").translate(str.maketrans("çğıöşüÇĞİÖŞÜıI", "cgiosuCGIOSUiI"))


def code_from_topic(topic, taken=()):
    words = []
    for word in fold_ascii(topic).split():
        letters = "".join(ch for ch in word if ch.isalnum())
        if letters:
            words.append(letters)
    if not words:
        return ""
    if len(words) == 1:
        base = words[0][:4].upper()
    else:
        base = "".join(word[0] for word in words)[:8].upper()
    if len(base) < 2:
        base = words[0][:4].upper()
    code = base[:8]
    number = 2
    used = {str(item).upper() for item in taken}
    while code in used:
        code = f"{base[:6]}{number}"[:8]
        number += 1
    return code


def note_list(raw):
    try:
        data = json.loads(raw or "[]")
    except json.JSONDecodeError:
        return []
    if not isinstance(data, list):
        return []
    notes = []
    for item in data:
        if not isinstance(item, dict):
            continue
        text = " ".join(str(item.get("note") or "").split())
        if text:
            notes.append({
                "status": str(item.get("status") or ""),
                "note": text[:1000],
                "at": str(item.get("at") or ""),
            })
    return notes[-30:]


class VkvLink:
    def read(self):
        path = kopru_path()
        try:
            lines = path.read_text(encoding="utf-8").splitlines()
        except OSError:
            return None
        if len(lines) < 2 or not lines[1].strip().isdigit():
            return None
        return {"token": lines[0].strip(), "port": int(lines[1].strip())}

    def call(self, path, payload=None):
        info = self.read()
        if info is None:
            return {"ok": False, "error": "VKV kapalı."}
        data = None if payload is None else json.dumps(payload, ensure_ascii=False).encode("utf-8")
        request = urllib.request.Request(
            f"http://127.0.0.1:{info['port']}{path}",
            data=data,
            method="GET" if data is None else "POST",
            headers={"Authorization": f"Bearer {info['token']}", "Content-Type": "application/json"},
        )
        try:
            with urllib.request.urlopen(request, timeout=6) as response:
                return json.loads(response.read().decode("utf-8"))
        except urllib.error.HTTPError as exc:
            try:
                return json.loads(exc.read().decode("utf-8"))
            except (OSError, json.JSONDecodeError):
                return {"ok": False, "error": "VKV isteği reddetti."}
        except (OSError, json.JSONDecodeError):
            return {"ok": False, "error": "VKV kapalı."}

    def healthy(self):
        return bool(self.call("/v1/durum").get("ok"))


class Click(tk.Frame):
    def __init__(self, master, text, command, bg, fg, hover=None, font=None, padx=14, pady=8, anchor="center"):
        super().__init__(master, bg=bg, cursor="hand2")
        self.command = command
        self._bg = bg
        self._fg = fg
        self._hover = hover or bg
        self.label = tk.Label(
            self, text=text, bg=bg, fg=fg, font=font or (FONT, 12, "bold"),
            padx=padx, pady=pady, cursor="hand2", anchor=anchor,
        )
        if anchor == "w":
            self.label.pack(fill="x")
        else:
            self.label.pack()
        for widget in (self, self.label):
            widget.bind("<Enter>", lambda _e: self._paint(self._hover))
            widget.bind("<Leave>", lambda _e: self._paint(self._bg))
            widget.bind("<Button-1>", self._click)

    def set_palette(self, bg, fg, hover=None):
        self._bg = bg
        self._fg = fg
        self._hover = hover or bg
        self.label.configure(fg=fg)
        self._paint(bg)

    def _paint(self, bg):
        self.configure(bg=bg)
        self.label.configure(bg=bg)

    def _click(self, _event):
        if self.command:
            self.command()


class TalepApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.configure(bg=CANVAS)
        self.minsize(980, 640)
        self.geometry("1100x720")
        try:
            self.store = Store()
        except ApiError as exc:
            messagebox.showerror(APP_NAME, str(exc))
            self.after(50, self.destroy)
            return
        self.link = VkvLink()
        self.user = None
        self.team_ids = set()
        self._ping = 0
        self.shell = tk.Frame(self, bg=CANVAS)
        self.shell.pack(fill="both", expand=True)
        self.protocol("WM_DELETE_WINDOW", self._close)
        self._route()

    def _close(self):
        self._ping += 1
        self._write_auth_log("Çıkış")
        self.store.close()
        self.destroy()

    def _route(self):
        for child in self.shell.winfo_children():
            child.destroy()
        if self.store.user_count() == 0:
            self._setup()
        else:
            self._login()

    def _card(self, title, note):
        card = tk.Frame(self.shell, bg=CARD, highlightbackground=LINE, highlightthickness=1, padx=22, pady=18)
        card.place(relx=0.5, rely=0.5, anchor="center")
        tk.Label(card, text=title, bg=CARD, fg=INK, font=(FONT, 20, "bold")).pack(anchor="w")
        tk.Label(card, text=note, bg=CARD, fg=MUTED, font=(FONT, 12), wraplength=420, justify="left").pack(anchor="w", pady=(4, 12))
        return card

    def _field(self, parent, label, secret=False):
        tk.Label(parent, text=label, bg=CARD, fg=MUTED, font=(FONT, 11, "bold")).pack(anchor="w", pady=(8, 2))
        var = tk.StringVar()
        entry = tk.Entry(parent, textvariable=var, relief="flat", highlightthickness=1, highlightbackground=LINE, font=(FONT, 13), show="•" if secret else "")
        entry.pack(fill="x", ipady=6)
        return var

    def _setup(self):
        card = self._card("İlk hesap", "Yönetici hesabı furkan.hamarat01@gmail.com olur. Sonraki kullanıcılar bir rolle açılır. Girişte e-posta, parola ve telefondaki 6 hane birlikte istenir.")
        name = self._field(card, "Ad soyad")
        email = self._field(card, "E-posta")
        email.set(ADMIN_EMAIL)
        password = self._field(card, "Parola", secret=True)
        again = self._field(card, "Parola tekrar", secret=True)
        secret = new_totp_secret()

        def save():
            if clean_email(email.get()) != ADMIN_EMAIL:
                messagebox.showwarning(APP_NAME, "Yönetici hesabı furkan.hamarat01@gmail.com olmalı.", parent=self)
                return
            if password.get() != again.get():
                messagebox.showwarning(APP_NAME, "Parolalar aynı değil.", parent=self)
                return
            try:
                self.store.add_user(name.get(), email.get(), password.get(), secret, True)
            except ValueError as exc:
                messagebox.showwarning(APP_NAME, str(exc), parent=self)
                return
            self.unbind("<Return>")
            self._show_secret(self.store.user_by_email(email.get()), secret)

        self._bind_return(save)
        Click(card, "Devam", save, ACCENT, "white", ACCENT_HOVER).pack(anchor="e", pady=(16, 0))

    def _show_secret(self, user, secret):
        note = app_dir() / "son-kurulum.txt"
        note.write_text(f"{user['email']}\n{secret}\n", encoding="utf-8")
        try:
            os.chmod(note, 0o600)
        except OSError:
            pass
        for child in self.shell.winfo_children():
            child.destroy()
        card = self._card(
            "Telefonuna ekle",
            "1. Google Authenticator veya Microsoft Authenticator aç.\n"
            "2. Hesap ekle, sonra Anahtarı gir.\n"
            "3. Aşağıdaki anahtarı yapıştır.\n"
            "4. Uygulamanın verdiği 6 haneyi yaz.",
        )
        row = tk.Frame(card, bg=CARD)
        row.pack(fill="x", pady=(0, 4))
        shown = tk.Entry(row, relief="flat", highlightthickness=1, highlightbackground=LINE, font=(FONT, 14), readonlybackground=CARD)
        shown.insert(0, secret)
        shown.configure(state="readonly")
        shown.pack(side="left", fill="x", expand=True, ipady=6)
        Click(row, "Kopyala", lambda: self._copy(secret), SOFT, ACCENT, HOVER).pack(side="left", padx=(8, 0))
        code = self._field(card, "6 haneli kod")

        def confirm():
            if not totp_ok(secret, code.get()):
                messagebox.showwarning(APP_NAME, "Kod uyuşmuyor. Telefondaki kodu, süresi dolmadan yaz.", parent=self)
                return
            try:
                note.unlink()
            except OSError:
                pass
            self.unbind("<Return>")
            self.user = self.store.user_by_email(user["email"])
            write_session(self.user)
            self._enter()

        self._bind_return(confirm)
        Click(card, "Kodu doğrula", confirm, ACCENT, "white", ACCENT_HOVER).pack(anchor="e", pady=(16, 0))

    def _login(self):
        card = self._card("Giriş", "Kendi e-posta adresin, parolan ve kimlik uygulamasındaki kod ile gir.")
        email = self._field(card, "E-posta")
        password = self._field(card, "Parola", secret=True)
        code = self._field(card, "Doğrulama kodu")

        def enter():
            device = device_snapshot()
            user = self.store.login(email.get(), password.get(), code.get())
            if user is None:
                self.store.add_login_log(email.get(), "Başarısız", device=device)
                messagebox.showwarning(APP_NAME, "E-posta, parola veya kod hatalı.", parent=self)
                return
            self.unbind("<Return>")
            changed = self._require_own_password(user)
            if changed is None:
                self._bind_return(enter)
                return
            self.user = changed
            write_session(changed)
            self.store.add_login_log(changed["email"], "Giriş", user=changed, device=device)
            self._enter()

        self._bind_return(enter)
        Click(card, "Giriş yap", enter, ACCENT, "white", ACCENT_HOVER).pack(anchor="e", pady=(16, 0))
        note = app_dir() / "son-kurulum.txt"
        if note.is_file():
            def resume():
                try:
                    saved = note.read_text(encoding="utf-8").splitlines()
                except OSError:
                    return
                if not saved:
                    return
                user = self.store.user_by_email(saved[0])
                if user is None:
                    return
                self.unbind("<Return>")
                self._show_secret(user, user["totp_secret"])

            Click(card, "Kurulum anahtarını göster", resume, CARD, ACCENT, HOVER).pack(anchor="e", pady=(8, 0))

    def _bind_return(self, command):
        self.bind("<Return>", lambda _e: command())

    def _copy(self, text):
        self.clipboard_clear()
        self.clipboard_append(text)
        self.update_idletasks()

    def _enter(self):
        self.unbind("<Return>")
        self.team_ids = self.store.team_ids(self.user["id"])
        self._ping += 1
        token = self._ping
        for child in self.shell.winfo_children():
            child.destroy()
        if not self.user["is_admin"]:
            self._member_notice()
            return
        sidebar = tk.Frame(self.shell, bg=NAV, width=220)
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)
        tk.Label(sidebar, text="Talep Yönetimi", bg=NAV, fg="white", font=(FONT, 16, "bold")).pack(anchor="w", padx=16, pady=(18, 2))
        tk.Label(sidebar, text=self.user["name"], bg=NAV, fg="white", font=(FONT, 12)).pack(anchor="w", padx=16)
        tk.Label(sidebar, text=self.user["email"], bg=NAV, fg=NAV_TEXT, font=(FONT, 11), wraplength=180, justify="left").pack(anchor="w", padx=16)
        role = self.user["role"] if "role" in self.user.keys() and self.user["role"] else ADMIN_ROLE
        tk.Label(sidebar, text=role, bg=NAV, fg="#8FDBA8", font=(FONT, 11, "bold")).pack(anchor="w", padx=16, pady=(2, 14))
        self._nav = {}
        self._nav_item(sidebar, "surec", "Süreç")
        self._nav_item(sidebar, "ekipler", "Ekipler")
        self._nav_item(sidebar, "kodlar", "Talep kodları")
        self._nav_item(sidebar, "kullanicilar", "Kullanıcılar")
        self._nav_item(sidebar, "kayitlar", "Giriş kayıtları")
        foot = tk.Frame(sidebar, bg=NAV)
        foot.pack(side="bottom", fill="x", pady=12)
        self.vkv_label = tk.Label(foot, text="VKV bakılıyor", bg=NAV, fg=NAV_TEXT, font=(FONT, 11), wraplength=180, justify="left")
        self.vkv_label.pack(anchor="w", padx=16, pady=(0, 8))
        Click(foot, "Çıkış", self._logout, NAV, NAV_TEXT, "#2E3A54", anchor="w", padx=16).pack(fill="x", padx=12)
        self.content = tk.Frame(self.shell, bg=CANVAS)
        self.content.pack(side="left", fill="both", expand=True)
        self.page = ""
        self._show_page("surec")
        self._ping_vkv(token)

    def _member_notice(self):
        card = self._card("Talebini VKV içinden aç", "Bu ekran talep sürecini yönetmek içindir. Yeni talebini VKV İş Takip içindeki Talepler sekmesinden aç.")
        Click(card, "Çıkış", self._logout, ACCENT, "white", ACCENT_HOVER).pack(anchor="e", pady=(8, 0))

    def _nav_item(self, parent, key, text):
        button = Click(parent, text, lambda: self._show_page(key), NAV, NAV_TEXT, "#2E3A54", anchor="w", padx=16, pady=10)
        button.pack(fill="x", padx=12, pady=2)
        self._nav[key] = button

    def _paint_nav(self):
        for key, button in self._nav.items():
            if key == self.page:
                button.set_palette("#2E3A54", "white", "#3A4866")
            else:
                button.set_palette(NAV, NAV_TEXT, "#2E3A54")

    def _show_page(self, name):
        self.page = name
        self.unbind("<Return>")
        for child in self.content.winfo_children():
            child.destroy()
        self._paint_nav()
        if name == "ekipler":
            self._build_teams()
        elif name == "kodlar":
            self._build_codes()
        elif name == "kullanicilar":
            self._build_users()
        elif name == "kayitlar":
            self._build_logs()
        else:
            self._build_process()

    def _build_process(self):
        head = tk.Frame(self.content, bg=CANVAS)
        head.pack(fill="x", padx=20, pady=(16, 8))
        tk.Label(head, text="Talep süreçleri", bg=CANVAS, fg=INK, font=(FONT, 22, "bold")).pack(side="left")
        Click(head, "Yenile", self._load_process, ACCENT, "white", ACCENT_HOVER).pack(side="right")
        tk.Label(self.content, text="Kullanıcı talebi VKV içinden açar. Aşamayı buradan ilerletirsin. Bekletme ve kapatma not ister.", bg=CANVAS, fg=MUTED, font=(FONT, 12), wraplength=760, justify="left").pack(anchor="w", padx=20)
        self.process_board = tk.Frame(self.content, bg=CANVAS)
        self.process_board.pack(fill="both", expand=True, padx=12, pady=(12, 16))
        self._process_rows = []
        self._load_process()

    def _load_process(self):
        if not hasattr(self, "process_board") or not self.process_board.winfo_exists():
            return
        result = self.link.call("/v1/talepler")
        for child in self.process_board.winfo_children():
            child.destroy()
        if not result.get("ok"):
            tk.Label(self.process_board, text=result.get("error") or "VKV kapalı. Süreçleri görmek için VKV'yi aç.", bg=CANVAS, fg=MUTED, font=(FONT, 13)).pack(anchor="w", padx=8, pady=12)
            return
        self._process_rows = list(result.get("talepler") or [])
        colors = {
            "Açık": ("#EAEEFD", "#3250E6"),
            "Üzerinde": ("#E7F3FF", "#1D4E89"),
            "Beklemede": ("#FCF3DC", "#B97A0A"),
            "Tamamlandı": ("#E5F4EA", "#1F8A4C"),
            "Reddedildi": ("#FBEAEB", "#D1343B"),
        }
        for status in STATUSES:
            rows = [row for row in self._process_rows if row.get("status") == status]
            tint, ink = colors.get(status, (CARD, INK))
            column = tk.Frame(self.process_board, bg=CANVAS)
            column.pack(side="left", fill="both", expand=True, padx=6)
            title = tk.Frame(column, bg=tint)
            title.pack(fill="x")
            tk.Label(title, text=f"{status}   {len(rows)}", bg=tint, fg=ink, font=(FONT, 12, "bold"), padx=10, pady=8).pack(anchor="w")
            body = tk.Frame(column, bg=CANVAS)
            body.pack(fill="both", expand=True, pady=(8, 0))
            if not rows:
                tk.Label(body, text="Kayıt yok", bg=CANVAS, fg=MUTED, font=(FONT, 11)).pack(anchor="w", padx=4)
            for row in rows:
                self._process_card(body, row)

    def _process_card(self, parent, row):
        card = tk.Frame(parent, bg=CARD, highlightbackground=LINE, highlightthickness=1, cursor="hand2")
        card.pack(fill="x", pady=4)
        widgets = [
            tk.Label(card, text=row.get("number") or "", bg=CARD, fg=ACCENT, font=(FONT, 11, "bold"), anchor="w", cursor="hand2"),
            tk.Label(card, text=row.get("title") or "", bg=CARD, fg=INK, font=(FONT, 12, "bold"), anchor="w", justify="left", wraplength=180, cursor="hand2"),
            tk.Label(card, text=f"{row.get('source_name') or '—'}  →  {row.get('target_name') or '—'}", bg=CARD, fg=MUTED, font=(FONT, 11), anchor="w", justify="left", wraplength=180, cursor="hand2"),
            tk.Label(card, text=f"{row.get('scope') or '—'}  ·  Atanan {row.get('assignee_name') or 'Havuz'}", bg=CARD, fg=INK, font=(FONT, 11), anchor="w", justify="left", wraplength=180, cursor="hand2"),
            tk.Label(card, text=f"Açılış {row.get('created_at') or '—'}", bg=CARD, fg=MUTED, font=(FONT, 11), anchor="w", cursor="hand2"),
        ]
        if row.get("request_code"):
            widgets.append(tk.Label(card, text=f"Talep kodu {row['request_code']}", bg=CARD, fg=MUTED, font=(FONT, 11), anchor="w", cursor="hand2"))
        if row.get("work_code"):
            widgets.append(tk.Label(card, text=f"İş kodu {row['work_code']}", bg=CARD, fg=MUTED, font=(FONT, 11), anchor="w", cursor="hand2"))
        for widget in widgets:
            widget.pack(fill="x", padx=10, pady=(6, 0))
        tk.Frame(card, bg=CARD, height=8).pack()

        def open_row(_event=None, item=row):
            self._process_detail(item)

        card.bind("<Button-1>", open_row)
        for widget in widgets:
            widget.bind("<Button-1>", open_row)

    def _process_detail(self, row):
        dialog = tk.Toplevel(self)
        dialog.title(row.get("number") or "Talep")
        dialog.configure(bg=CANVAS)
        dialog.transient(self)
        dialog.resizable(False, False)
        card = tk.Frame(dialog, bg=CARD, padx=16, pady=14, highlightbackground=LINE, highlightthickness=1)
        card.pack(padx=12, pady=12)
        tk.Label(card, text=f"{row.get('number') or ''}   ·   {row.get('title') or ''}", bg=CARD, fg=INK, font=(FONT, 15, "bold"), wraplength=460, justify="left").pack(anchor="w")
        meta = f"{row.get('source_name') or '—'}  →  {row.get('target_name') or '—'}"
        if row.get("author_name"):
            meta += f"   ·   Açan {row['author_name']}"
        if row.get("scope"):
            meta += f"   ·   {row['scope']}"
        meta += f"   ·   Atanan {row.get('assignee_name') or 'Havuz'}"
        if row.get("created_at"):
            meta += f"   ·   Açılış {row['created_at']}"
        if row.get("due_date"):
            meta += f"   ·   İstenen {row['due_date']}"
        if row.get("request_code"):
            meta += f"   ·   Talep kodu {row['request_code']}"
        if row.get("work_code"):
            meta += f"   ·   İş kodu {row['work_code']}"
        if row.get("priority"):
            meta += f"   ·   {row['priority']}"
        tk.Label(card, text=meta, bg=CARD, fg=MUTED, font=(FONT, 11), wraplength=460, justify="left").pack(anchor="w", pady=(4, 8))
        tk.Label(card, text=row.get("description") or "Açıklama yok", bg=CARD, fg=INK, font=(FONT, 12), wraplength=460, justify="left").pack(anchor="w")
        if row.get("file_name"):
            tk.Label(card, text=f"Ek: {row['file_name']}", bg=CARD, fg=MUTED, font=(FONT, 11)).pack(anchor="w", pady=(6, 0))
        for item in (row.get("notes") or [])[-4:]:
            when = f"{item.get('status') or ''}  ·  {item.get('at') or ''}".strip(" ·")
            tk.Label(card, text=when, bg=CARD, fg=MUTED, font=(FONT, 11, "bold")).pack(anchor="w", pady=(8, 0))
            tk.Label(card, text=item.get("note") or "", bg=CARD, fg=INK, font=(FONT, 12), wraplength=460, justify="left").pack(anchor="w")
        actions = tk.Frame(card, bg=CARD)
        actions.pack(fill="x", pady=(14, 0))
        for status in STATUSES:
            if status == row.get("status"):
                continue
            Click(actions, status, lambda value=status: self._move_process(dialog, row, value), SOFT, ACCENT, HOVER).pack(side="left", padx=(0, 6))

    def _move_process(self, dialog, row, status):
        def work():
            note = ""
            if status in NOTE_STATUSES:
                note = self._ask_note(status)
                if not note:
                    return
            result = self.link.call("/v1/gorev/durum", {"id": row.get("id"), "status": status, "note": note})
            if not result.get("ok"):
                parent = dialog if dialog.winfo_exists() else self
                messagebox.showwarning(APP_NAME, result.get("error") or "Durum kaydedilemedi.", parent=parent)
                return
            if dialog.winfo_exists():
                dialog.destroy()
            self._load_process()

        self.after(20, work)

    def _build_requests(self):
        head = tk.Frame(self.content, bg=CANVAS)
        head.pack(fill="x", padx=20, pady=(16, 8))
        if self.user["is_admin"]:
            page_title, page_note, button_text = "Tüm talepler", "Bütün ekiplerin açtığı ve aldığı talepler.", "Talep aç"
        else:
            page_title, page_note, button_text = "Taleplerim", "Ekibinin iç talepleri ve ekibine gelen dış talepler.", "Yeni talep"
        tk.Label(head, text=page_title, bg=CANVAS, fg=INK, font=(FONT, 22, "bold")).pack(side="left")
        Click(head, button_text, self._new_request, ACCENT, "white", ACCENT_HOVER).pack(side="right")
        tk.Label(self.content, text=page_note, bg=CANVAS, fg=MUTED, font=(FONT, 12)).pack(anchor="w", padx=20)
        table = tk.Frame(self.content, bg=CARD, highlightbackground=LINE, highlightthickness=1)
        table.pack(fill="both", expand=True, padx=20, pady=(10, 8))
        if self.user["is_admin"]:
            columns = ("number", "title", "source", "target", "work", "status")
            labels = (("number", "Talep no", 110), ("title", "Başlık", 220), ("source", "Kimden", 140), ("target", "Kime", 140), ("work", "İş kodu", 110), ("status", "Durum", 120))
        else:
            columns = ("number", "title", "work", "target", "status")
            labels = (("number", "Talep no", 120), ("title", "Başlık", 260), ("work", "İş kodu", 120), ("target", "Hedef", 160), ("status", "Durum", 130))
        self.tree = ttk.Treeview(table, columns=columns, show="headings")
        for key, text, width in labels:
            self.tree.heading(key, text=text)
            self.tree.column(key, width=width, anchor="w")
        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<<TreeviewSelect>>", lambda _e: self._show_selected())
        self.detail = tk.Frame(self.content, bg=CARD, highlightbackground=LINE, highlightthickness=1, height=280)
        self.detail.pack(fill="x", padx=20, pady=(0, 16))
        self.detail.pack_propagate(False)
        self._reload()

    def _logout(self):
        self._write_auth_log("Çıkış")
        clear_session()
        self.user = None
        self.team_ids = set()
        self._ping += 1
        self._route()

    def _write_auth_log(self, event):
        if self.user is None:
            return
        try:
            self.store.add_login_log(self.user["email"], event, user=self.user, device=device_snapshot())
        except (ValueError, ApiError, OSError):
            pass

    def _ping_vkv(self, token):
        if token != self._ping or not self.winfo_exists():
            return

        def work():
            ok = self.link.healthy()

            def apply():
                if token != self._ping or not self.winfo_exists():
                    return
                if hasattr(self, "vkv_label") and self.vkv_label.winfo_exists():
                    self.vkv_label.configure(
                        text="VKV bağlı" if ok else "VKV kapalı. Talep burada kalır.",
                        fg="#8FDBA8" if ok else NAV_TEXT,
                    )
                if ok and self.user is not None:
                    self._flush_vkv()
                self.after(4000, lambda: self._ping_vkv(token))

            try:
                self.after(0, apply)
            except tk.TclError:
                pass

        threading.Thread(target=work, daemon=True).start()

    def _reload(self, focus=None):
        if not hasattr(self, "tree") or not self.tree.winfo_exists():
            return
        for item in self.tree.get_children():
            self.tree.delete(item)
        self._rows = {}
        for row in self.store.visible_requests(self.user):
            if self.user["is_admin"]:
                values = (row["number"], row["title"], row["source_name"], row["target_name"], row["work_code"], row["status"])
            else:
                values = (row["number"], row["title"], row["work_code"], row["target_name"], row["status"])
            self.tree.insert("", "end", iid=str(row["id"]), values=values)
            self._rows[row["id"]] = row
        if focus and self.tree.exists(str(focus)):
            self.tree.selection_set(str(focus))
            self.tree.focus(str(focus))
        self._show_selected()

    def _selected(self):
        if not hasattr(self, "tree") or not self.tree.winfo_exists():
            return None
        picked = self.tree.selection()
        if not picked:
            return None
        return self._rows.get(int(picked[0]))

    def _show_selected(self):
        if not hasattr(self, "detail") or not self.detail.winfo_exists():
            return
        for child in self.detail.winfo_children():
            child.destroy()
        row = self._selected()
        if row is None:
            tk.Label(self.detail, text="Bir talep seç.", bg=CARD, fg=MUTED, font=(FONT, 12)).pack(anchor="w", padx=12, pady=12)
            return
        actions = tk.Frame(self.detail, bg=CARD)
        actions.pack(side="bottom", fill="x", padx=12, pady=(0, 10))
        if row["file_path"]:
            Click(actions, "Dosyayı aç", lambda path=row["file_path"]: self._open_file(path), SOFT, ACCENT, HOVER).pack(side="left")
        if self.user["is_admin"] or row["target_team_id"] in self.team_ids:
            status = tk.StringVar(value=row["status"])
            ttk.Combobox(actions, textvariable=status, values=list(STATUSES), state="readonly", width=16, font=(FONT, 12)).pack(side="left", padx=(8, 0))
            Click(actions, "Durumu kaydet", lambda rid=row["id"], var=status: self._set_status(rid, var.get()), ACCENT, "white", ACCENT_HOVER).pack(side="left", padx=(8, 0))
        tk.Label(self.detail, text=f"{row['number']}  ·  {row['title']}", bg=CARD, fg=INK, font=(FONT, 14, "bold")).pack(anchor="w", padx=12, pady=(10, 2))
        meta = f"{row['source_name']} → {row['target_name']}   ·   İş kodu {row['work_code']}   ·   {row['author_name']}"
        if row["due_date"]:
            meta += f"   ·   {row['due_date']}"
        tk.Label(self.detail, text=meta, bg=CARD, fg=MUTED, font=(FONT, 11)).pack(anchor="w", padx=12)
        tk.Label(self.detail, text=row["description"] or "Açıklama yok", bg=CARD, fg=INK, font=(FONT, 12), wraplength=760, justify="left").pack(anchor="w", padx=12, pady=(6, 4))
        saved_notes = row["notes_json"] if "notes_json" in row.keys() else ""
        for item in note_list(saved_notes)[-3:]:
            when = f"{item['status']}  ·  {item['at']}".strip(" ·")
            tk.Label(self.detail, text=when, bg=CARD, fg=MUTED, font=(FONT, 11, "bold")).pack(anchor="w", padx=12, pady=(4, 0))
            tk.Label(self.detail, text=item["note"], bg=CARD, fg=INK, font=(FONT, 12), wraplength=760, justify="left").pack(anchor="w", padx=12)

    def _open_file(self, path):
        file = Path(path)
        if not file.is_file() or file.is_symlink():
            messagebox.showwarning(APP_NAME, "Dosya bulunamadı.", parent=self)
            return
        if sys.platform == "darwin":
            subprocess.Popen(["open", str(file)])
        elif sys.platform == "win32":
            os.startfile(file)
        else:
            subprocess.Popen(["xdg-open", str(file)])

    def _ask_note(self, status):
        closing = status in ("Tamamlandı", "Reddedildi")
        dialog = tk.Toplevel(self)
        dialog.title("Talebi kapat" if closing else "Beklemeye al")
        dialog.configure(bg=CANVAS)
        dialog.transient(self)
        dialog.resizable(False, False)
        card = tk.Frame(dialog, bg=CARD, padx=18, pady=16, highlightbackground=LINE, highlightthickness=1)
        card.pack(padx=12, pady=12)
        tk.Label(card, text="Kapatma notu" if closing else "Bekleme notu", bg=CARD, fg=INK, font=(FONT, 16, "bold")).pack(anchor="w")
        prompt = "Talebi neden kapattığını yaz." if closing else "Talebi neden beklemeye aldığını yaz."
        tk.Label(card, text=prompt, bg=CARD, fg=MUTED, font=(FONT, 12), wraplength=420, justify="left").pack(anchor="w", pady=(4, 10))
        box = tk.Text(
            card, width=48, height=5, wrap="word", font=(FONT, 13),
            bg="#FFFFFF", fg=INK, insertbackground=INK, insertwidth=2,
            relief="flat", highlightthickness=1, highlightbackground=LINE, highlightcolor=ACCENT,
            padx=8, pady=6, takefocus=True,
        )
        box.pack(fill="x")
        result = {"note": ""}

        def save():
            text = " ".join(box.get("1.0", "end").split())
            if not text:
                messagebox.showwarning(APP_NAME, "Not yaz.", parent=dialog)
                return
            result["note"] = text[:1000]
            dialog.destroy()

        actions = tk.Frame(card, bg=CARD)
        actions.pack(fill="x", pady=(12, 0))
        def focus_box():
            if not dialog.winfo_exists():
                return
            dialog.lift()
            dialog.focus_force()
            box.focus_force()

        Click(actions, "Vazgeç", dialog.destroy, CARD, INK, HOVER).pack(side="left")
        Click(actions, "Notu kaydet", save, ACCENT, "white", ACCENT_HOVER).pack(side="right")
        dialog.bind("<Escape>", lambda _e: dialog.destroy())
        box.bind("<Button-1>", lambda _e: box.focus_set())
        dialog.update_idletasks()
        dialog.grab_set()
        dialog.after(30, focus_box)
        self.wait_window(dialog)
        return result["note"]

    def _set_status(self, request_id, status):
        row = self.store.request(request_id)
        if row is None or row["status"] == status:
            return
        note = ""
        if status in NOTE_STATUSES:
            note = self._ask_note(status)
            if not note:
                self._reload(request_id)
                return
        try:
            self.store.set_status(request_id, status, note)
        except ValueError as exc:
            messagebox.showwarning(APP_NAME, str(exc), parent=self)
            return
        if row["vkv_task_id"]:
            self.link.call("/v1/gorev/durum", {"id": row["vkv_task_id"], "status": status, "note": note})
        self._reload(request_id)

    def _new_request(self):
        if self.user["is_admin"]:
            self._admin_request_dialog()
        else:
            self._user_request_dialog()

    def _request_dialog(self, title, note):
        dialog = tk.Toplevel(self)
        dialog.title(title)
        dialog.configure(bg=CANVAS)
        dialog.transient(self)
        dialog.resizable(False, False)
        card = tk.Frame(dialog, bg=CARD, padx=16, pady=14, highlightbackground=LINE, highlightthickness=1)
        card.pack(padx=12, pady=12)
        tk.Label(card, text=title, bg=CARD, fg=INK, font=(FONT, 16, "bold")).grid(row=0, column=0, columnspan=2, sticky="w")
        tk.Label(card, text=note, bg=CARD, fg=MUTED, font=(FONT, 11), wraplength=460, justify="left").grid(row=1, column=0, columnspan=2, sticky="w", pady=(4, 8))
        return dialog, card

    def _form_row(self, card, row, label, var, choices=None):
        tk.Label(card, text=label, bg=CARD, fg=MUTED, font=(FONT, 11, "bold")).grid(row=row, column=0, sticky="w", pady=6, padx=(0, 10))
        if choices is None:
            widget = tk.Entry(card, textvariable=var, relief="flat", highlightthickness=1, highlightbackground=LINE, font=(FONT, 12), width=44)
        else:
            widget = ttk.Combobox(card, textvariable=var, values=choices, state="readonly", width=42, font=(FONT, 12))
        widget.grid(row=row, column=1, sticky="ew", pady=6, ipady=4)
        return widget

    def _common_request_fields(self, card, start):
        work = tk.StringVar()
        title = tk.StringVar()
        priority = tk.StringVar(value="Normal")
        due = tk.StringVar()
        file_var = tk.StringVar()
        self._form_row(card, start, "İş kodu", work)
        self._form_row(card, start + 1, "Başlık", title)
        self._form_row(card, start + 2, "Öncelik", priority, list(PRIORITIES))
        self._form_row(card, start + 3, "İstenen tarih", due)
        tk.Label(card, text="Açıklama", bg=CARD, fg=MUTED, font=(FONT, 11, "bold")).grid(row=start + 4, column=0, sticky="nw", pady=6, padx=(0, 10))
        description = tk.Text(card, width=44, height=5, wrap="word", font=(FONT, 12), relief="flat", highlightthickness=1, highlightbackground=LINE)
        description.grid(row=start + 4, column=1, sticky="ew", pady=6)
        file_row = tk.Frame(card, bg=CARD)
        file_row.grid(row=start + 5, column=1, sticky="w", pady=4)
        tk.Label(file_row, textvariable=file_var, bg=CARD, fg=MUTED, font=(FONT, 11)).pack(side="left")

        def choose_file():
            picked = filedialog.askopenfilename(parent=card.winfo_toplevel(), filetypes=[("PDF ve görüntü", "*.pdf *.png *.jpg *.jpeg")])
            if picked:
                file_var.set(Path(picked).name)
                file_var._path = picked

        Click(card, "Dosya ekle", choose_file, SOFT, ACCENT, HOVER).grid(row=start + 5, column=0, sticky="w")
        return work, title, priority, due, description, file_var, start + 6

    def _commit_request(self, dialog, source_team, target_team, work, title, description, priority, due, file_var):
        if source_team is None or target_team is None:
            messagebox.showwarning(APP_NAME, "Ekip seç.", parent=dialog)
            return
        try:
            row = self.store.create_request(
                self.user["id"], source_team["id"], target_team["id"], work.get(), title.get(),
                description.get("1.0", "end"), priority.get(), due.get(),
            )
        except ValueError as exc:
            messagebox.showwarning(APP_NAME, str(exc), parent=dialog)
            return
        file_error = ""
        picked = getattr(file_var, "_path", "")
        if picked:
            try:
                self.store.attach(row["id"], picked)
                row = self.store.request(row["id"])
            except (ValueError, OSError) as exc:
                file_error = str(exc)
        result = self._send_vkv(row)
        if result.get("ok"):
            self.store.mark_vkv(row["id"], result.get("id"), result.get("issue_key"))
        dialog.destroy()
        self._reload(row["id"])
        notes = []
        if file_error:
            notes.append(f"Dosya eklenemedi: {file_error}")
        if not result.get("ok"):
            notes.append(f"VKV'ye iletilemedi: {result.get('error') or 'bağlantı yok.'}")
        if notes:
            messagebox.showinfo(APP_NAME, f"Talep {row['number']} kaydedildi. " + " ".join(notes), parent=self)

    def _user_request_dialog(self):
        mine = [team for team in self.store.teams() if team["id"] in self.team_ids]
        if not mine:
            messagebox.showinfo(APP_NAME, "Önce bir ekibe girmen gerekir. Yönetici seni Ekipler ekranından ekler.", parent=self)
            return
        dialog, card = self._request_dialog("Yeni talep", "Kendi ekibine iç talep açarsın. Başka ekibe gidecekse onun dış talep kodunu seçersin.")
        source = tk.StringVar(value=self._team_label(mine[0]))
        inside = "Ekibimin içinde"
        target = tk.StringVar(value=inside)
        if len(mine) == 1:
            tk.Label(card, text="Ekibim", bg=CARD, fg=MUTED, font=(FONT, 11, "bold")).grid(row=2, column=0, sticky="w", pady=6, padx=(0, 10))
            tk.Label(card, text=self._team_label(mine[0]), bg=CARD, fg=INK, font=(FONT, 13)).grid(row=2, column=1, sticky="w", pady=6)
        else:
            self._form_row(card, 2, "Ekibim", source, [self._team_label(team) for team in mine])
        hedef = self._form_row(card, 3, "Hedef", target, [inside])

        def paint_targets(*_args):
            team = self._team_from_label(source.get(), mine)
            others = [self._team_label(item) for item in self.store.teams() if team is None or item["id"] != team["id"]]
            choices = [inside] + others
            hedef.configure(values=choices)
            if target.get() not in choices:
                target.set(inside)

        source.trace_add("write", paint_targets)
        paint_targets()
        work, title, priority, due, description, file_var, last = self._common_request_fields(card, 4)

        def save():
            source_team = self._team_from_label(source.get(), mine)
            if target.get() == inside:
                target_team = source_team
            else:
                target_team = self._team_from_label(target.get(), self.store.teams())
            self._commit_request(dialog, source_team, target_team, work, title, description, priority, due, file_var)

        Click(card, "Kaydet", save, ACCENT, "white", ACCENT_HOVER).grid(row=last, column=1, sticky="e", pady=(10, 0))
        dialog.bind("<Escape>", lambda _e: dialog.destroy())
        dialog.grab_set()

    def _admin_request_dialog(self):
        teams = list(self.store.teams())
        if not teams:
            messagebox.showinfo(APP_NAME, "Önce Ekipler sayfasından ekip ekle.", parent=self)
            return
        dialog, card = self._request_dialog("Yönetici talebi", "İstediğin ekipten istediğin ekibe talep açarsın. Aynı ekibi iki yana seçersen talep o ekibin içinde kalır.")
        source = tk.StringVar(value=self._team_label(teams[0]))
        target = tk.StringVar(value=self._team_label(teams[0]))
        self._form_row(card, 2, "Kimden", source, [self._team_label(team) for team in teams])
        self._form_row(card, 3, "Kime", target, [self._team_label(team) for team in teams])
        work, title, priority, due, description, file_var, last = self._common_request_fields(card, 4)

        def save():
            source_team = self._team_from_label(source.get(), teams)
            target_team = self._team_from_label(target.get(), teams)
            self._commit_request(dialog, source_team, target_team, work, title, description, priority, due, file_var)

        Click(card, "Kaydet", save, ACCENT, "white", ACCENT_HOVER).grid(row=last, column=1, sticky="e", pady=(10, 0))
        dialog.bind("<Escape>", lambda _e: dialog.destroy())
        dialog.grab_set()

    def _vkv_payload(self, row):
        return {
            "talep_no": row["number"],
            "title": row["title"],
            "description": row["description"] or "",
            "work_code": row["work_code"] or "",
            "source_name": row["source_name"] or "",
            "target_name": row["target_name"] or "",
            "author_name": row["author_name"] or "",
            "author_email": row["author_email"] or "",
            "priority": row["priority"] or "Normal",
            "due_date": row["due_date"] or "",
            "status": row["status"] or "Açık",
            "file_name": Path(row["file_path"]).name if row["file_path"] else "",
        }

    def _send_vkv(self, row):
        return self.link.call("/v1/gorev", self._vkv_payload(row))

    def _flush_vkv(self):
        changed = False
        for row in self.store.pending_vkv():
            result = self._send_vkv(row)
            if not result.get("ok"):
                continue
            self.store.mark_vkv(row["id"], result.get("id"), result.get("issue_key"))
            changed = True
        if changed and self.page == "talepler" and hasattr(self, "tree") and self.tree.winfo_exists():
            self._reload()

    def _team_label(self, team):
        return f"{team['code']}  ·  {team['name']}"

    def _team_from_label(self, label, teams):
        for team in teams:
            if self._team_label(team) == label:
                return team
        return None

    def _page_head(self, title, note):
        tk.Label(self.content, text=title, bg=CANVAS, fg=INK, font=(FONT, 22, "bold")).pack(anchor="w", padx=20, pady=(16, 2))
        tk.Label(self.content, text=note, bg=CANVAS, fg=MUTED, font=(FONT, 12), wraplength=720, justify="left").pack(anchor="w", padx=20)

    def _form_card(self):
        card = tk.Frame(self.content, bg=CARD, padx=16, pady=14, highlightbackground=LINE, highlightthickness=1)
        card.pack(fill="x", padx=20, pady=(12, 8))
        return card

    def _line(self, parent, label, secret=False):
        tk.Label(parent, text=label, bg=CARD, fg=MUTED, font=(FONT, 11, "bold")).pack(anchor="w", pady=(8, 2))
        var = tk.StringVar()
        tk.Entry(
            parent, textvariable=var, relief="flat", highlightthickness=1, highlightbackground=LINE,
            font=(FONT, 13), show="•" if secret else "",
        ).pack(fill="x", ipady=5)
        return var

    def _table(self, columns):
        frame = tk.Frame(self.content, bg=CARD, highlightbackground=LINE, highlightthickness=1)
        frame.pack(fill="both", expand=True, padx=20, pady=(4, 16))
        tree = ttk.Treeview(frame, columns=tuple(key for key, _text, _width in columns), show="headings")
        for key, text, width in columns:
            tree.heading(key, text=text)
            tree.column(key, width=width, anchor="w")
        tree.pack(fill="both", expand=True)
        return tree

    def _require_own_password(self, user):
        if "must_change" not in user.keys() or not user["must_change"]:
            return user
        dialog = tk.Toplevel(self)
        dialog.title("Parola belirle")
        dialog.configure(bg=CANVAS)
        dialog.transient(self)
        dialog.resizable(False, False)
        card = tk.Frame(dialog, bg=CARD, padx=16, pady=14, highlightbackground=LINE, highlightthickness=1)
        card.pack(padx=12, pady=12)
        tk.Label(card, text="Parolanı belirle", bg=CARD, fg=INK, font=(FONT, 16, "bold")).pack(anchor="w")
        tk.Label(
            card, text="İlk girişin. Yönetici parolasını bir kez kullandın. Bundan sonra kendi parolan geçerli.",
            bg=CARD, fg=MUTED, font=(FONT, 12), wraplength=380, justify="left",
        ).pack(anchor="w", pady=(4, 8))
        first = tk.StringVar()
        second = tk.StringVar()
        done = {"user": None}
        for label, var in (("Yeni parola", first), ("Yeni parola tekrar", second)):
            tk.Label(card, text=label, bg=CARD, fg=MUTED, font=(FONT, 11, "bold")).pack(anchor="w", pady=(8, 2))
            tk.Entry(card, textvariable=var, show="•", relief="flat", highlightthickness=1, highlightbackground=LINE, font=(FONT, 13), width=36).pack(fill="x", ipady=5)

        def save():
            if len(first.get()) < 6:
                messagebox.showwarning(APP_NAME, "Parola en az 6 karakter olmalı.", parent=dialog)
                return
            if first.get() != second.get():
                messagebox.showwarning(APP_NAME, "Parolalar aynı değil.", parent=dialog)
                return
            try:
                self.store.set_password(user["id"], first.get(), False)
            except ValueError as exc:
                messagebox.showwarning(APP_NAME, str(exc), parent=dialog)
                return
            done["user"] = self.store.user_by_email(user["email"])
            dialog.destroy()

        Click(card, "Parolayı kaydet", save, ACCENT, "white", ACCENT_HOVER).pack(anchor="e", pady=(14, 0))
        dialog.bind("<Return>", lambda _e: save())
        dialog.bind("<Escape>", lambda _e: dialog.destroy())
        dialog.grab_set()
        self.wait_window(dialog)
        return done["user"]

    def _build_teams(self):
        if not self.user["is_admin"]:
            self._show_page("talepler")
            return
        self._page_head("Ekipler", "Ekip adı, konu ve dış talep kodu. Kod, yazdığın konudan kısaltılır. Satırın sağındaki kalemle sonra da düzenlersin.")
        card = self._form_card()
        team_name = self._line(card, "Ekip adı")
        topic = self._line(card, "Konu")
        code = tk.StringVar()
        tk.Label(card, text="Dış talep kodu", bg=CARD, fg=MUTED, font=(FONT, 11, "bold")).pack(anchor="w", pady=(8, 2))
        shown = tk.Entry(card, textvariable=code, relief="flat", highlightthickness=1, highlightbackground=LINE, font=(FONT, 13), state="readonly", readonlybackground=CARD)
        shown.pack(fill="x", ipady=5)

        def refresh_code(*_args):
            taken = [team["code"] for team in self.store.teams()]
            code.set(code_from_topic(topic.get(), taken))

        topic.trace_add("write", refresh_code)

        def add_team():
            try:
                self.store.add_team(team_name.get(), code.get(), topic.get())
            except ValueError as exc:
                messagebox.showwarning(APP_NAME, str(exc), parent=self)
                return
            team_name.set("")
            topic.set("")
            code.set("")
            paint()

        Click(card, "Ekip ekle", add_team, ACCENT, "white", ACCENT_HOVER).pack(anchor="e", pady=(12, 0))
        listing = tk.Frame(self.content, bg=CANVAS)
        listing.pack(fill="both", expand=True, padx=20, pady=(4, 16))

        def paint():
            for child in listing.winfo_children():
                child.destroy()
            head = tk.Frame(listing, bg=CANVAS)
            head.pack(fill="x", pady=(0, 4))
            for text, width in (("Ekip adı", 22), ("Konu", 24), ("Dış talep kodu", 16), ("Kullanıcı", 10)):
                tk.Label(head, text=text, bg=CANVAS, fg=MUTED, font=(FONT, 11, "bold"), width=width, anchor="w").pack(side="left")
            for team in self.store.teams():
                row = tk.Frame(listing, bg=CARD, highlightbackground=LINE, highlightthickness=1)
                row.pack(fill="x", pady=2)
                topic_text = team["topic"] if "topic" in team.keys() and team["topic"] else "—"
                for value, width in (
                    (team["name"], 22),
                    (topic_text, 24),
                    (team["code"], 16),
                    (str(self.store.member_count(team["id"])), 10),
                ):
                    tk.Label(row, text=value, bg=CARD, fg=INK, font=(FONT, 12), width=width, anchor="w").pack(side="left", padx=(8, 0), pady=6)
                Click(row, "✎", lambda item=team: self._edit_team(item, paint), CARD, ACCENT, HOVER, padx=10, pady=4).pack(side="right", padx=6)

        paint()

    def _edit_team(self, team, refresh):
        dialog = tk.Toplevel(self)
        dialog.title("Ekibi düzenle")
        dialog.configure(bg=CANVAS)
        dialog.transient(self)
        dialog.resizable(False, False)
        card = tk.Frame(dialog, bg=CARD, padx=16, pady=14, highlightbackground=LINE, highlightthickness=1)
        card.pack(padx=12, pady=12)
        tk.Label(card, text="Ekibi düzenle", bg=CARD, fg=INK, font=(FONT, 16, "bold")).pack(anchor="w")
        name = tk.StringVar(value=team["name"])
        topic = tk.StringVar(value=team["topic"] if "topic" in team.keys() and team["topic"] else "")
        current = tk.StringVar(value=team["code"])
        manual = {"on": False}
        for label, var in (("Ekip adı", name), ("Konu", topic), ("Dış talep kodu", current)):
            tk.Label(card, text=label, bg=CARD, fg=MUTED, font=(FONT, 11, "bold")).pack(anchor="w", pady=(8, 2))
            tk.Entry(card, textvariable=var, relief="flat", highlightthickness=1, highlightbackground=LINE, font=(FONT, 13), width=36).pack(fill="x", ipady=5)

        def follow(*_args):
            if manual["on"]:
                return
            taken = [item["code"] for item in self.store.teams() if item["id"] != team["id"]]
            current.set(code_from_topic(topic.get(), taken))

        topic.trace_add("write", follow)

        def save():
            manual["on"] = True
            try:
                self.store.update_team(team["id"], name.get(), topic.get(), current.get())
            except ValueError as exc:
                manual["on"] = False
                messagebox.showwarning(APP_NAME, str(exc), parent=dialog)
                return
            dialog.destroy()
            refresh()

        Click(card, "Kaydet", save, ACCENT, "white", ACCENT_HOVER).pack(anchor="e", pady=(14, 0))
        dialog.bind("<Escape>", lambda _e: dialog.destroy())
        dialog.grab_set()

    def _build_codes(self):
        if not self.user["is_admin"]:
            self._show_page("surec")
            return
        self._page_head("Talep kodları", "İş kodunu ve ekibi burada belirlersin. Kullanıcı VKV'de bu kodu seçer, açılmasını istediği başlığı yazar.")
        teams = list(self.store.teams())
        card = self._form_card()
        if not teams:
            tk.Label(card, text="Önce ekip ekle.", bg=CARD, fg=INK, font=(FONT, 13)).pack(anchor="w")
            Click(card, "Ekiplere git", lambda: self._show_page("ekipler"), ACCENT, "white", ACCENT_HOVER).pack(anchor="e", pady=(12, 0))
        else:
            code = self._line(card, "Talep kodu")
            tk.Label(card, text="Örnek: BT-03. Kullanıcı bu kodu listeden seçer.", bg=CARD, fg=MUTED, font=(FONT, 11)).pack(anchor="w", pady=(4, 0))
            heading = self._line(card, "Başlık")
            tk.Label(card, text="Kodun ne olduğu buradan anlaşılır. Örnek: Bilgisayar kurulumu.", bg=CARD, fg=MUTED, font=(FONT, 11)).pack(anchor="w", pady=(4, 0))
            tk.Label(card, text="Ekip", bg=CARD, fg=MUTED, font=(FONT, 11, "bold")).pack(anchor="w", pady=(8, 2))
            membership = tk.StringVar(value=self._team_label(teams[0]))
            ttk.Combobox(
                card, textvariable=membership, values=["Hepsi"] + [self._team_label(team) for team in teams],
                state="readonly", font=(FONT, 12),
            ).pack(fill="x", ipady=4)
            target_wrap = tk.Frame(card, bg=CARD)
            tk.Label(target_wrap, text="Hedef ekip", bg=CARD, fg=MUTED, font=(FONT, 11, "bold")).pack(anchor="w", pady=(8, 2))
            target_var = tk.StringVar(value=self._team_label(teams[0]))
            ttk.Combobox(
                target_wrap, textvariable=target_var, values=[self._team_label(team) for team in teams],
                state="readonly", font=(FONT, 12),
            ).pack(fill="x", ipady=4)
            tk.Label(target_wrap, text="Kod bu ekibe düşer. Her ekip bu kodu kendi listesinde görür.", bg=CARD, fg=MUTED, font=(FONT, 11)).pack(anchor="w", pady=(4, 0))
            scope_label = tk.Label(card, text="Tür", bg=CARD, fg=MUTED, font=(FONT, 11, "bold"))
            scope_label.pack(anchor="w", pady=(8, 2))

            def paint_share(*_args):
                if membership.get() == "Hepsi":
                    target_wrap.pack(fill="x", before=scope_label)
                else:
                    target_wrap.pack_forget()

            membership.trace_add("write", paint_share)
            scope_var = tk.StringVar(value="İç")
            scope_frame = tk.Frame(card, bg=CARD)
            scope_frame.pack(anchor="w")
            tk.Radiobutton(scope_frame, text="İç", variable=scope_var, value="İç", bg=CARD, fg=INK, font=(FONT, 12), selectcolor=CARD, activebackground=CARD).pack(side="left", padx=(0, 16))
            tk.Radiobutton(scope_frame, text="Dış", variable=scope_var, value="Dış", bg=CARD, fg=INK, font=(FONT, 12), selectcolor=CARD, activebackground=CARD).pack(side="left")
            work = self._line(card, "İş kodu")
            syncing = {"on": False}
            locked = {"on": False}

            def follow(*_args):
                if locked["on"]:
                    return
                syncing["on"] = True
                work.set(clean_request_code(code.get()))
                syncing["on"] = False

            def edited(*_args):
                if not syncing["on"]:
                    locked["on"] = True

            code.trace_add("write", follow)
            work.trace_add("write", edited)

            def add_code():
                shared = membership.get() == "Hepsi"
                picked = target_var.get() if shared else membership.get()
                team = self._team_from_label(picked, self.store.teams())
                if team is None:
                    messagebox.showwarning(APP_NAME, "Hedef ekip seç." if shared else "Ekip seç.", parent=self)
                    return
                try:
                    self.store.add_request_code(code.get(), team["id"], work.get(), heading.get(), scope_var.get(), shared)
                except ValueError as exc:
                    messagebox.showwarning(APP_NAME, str(exc), parent=self)
                    return
                code.set("")
                heading.set("")
                work.set("")
                scope_var.set("İç")
                locked["on"] = False
                paint()

            Click(card, "Talep kodu ekle", add_code, ACCENT, "white", ACCENT_HOVER).pack(anchor="e", pady=(12, 0))
        filter_frame = tk.Frame(self.content, bg=CANVAS)
        filter_frame.pack(fill="x", padx=20, pady=(8, 0))
        tk.Label(filter_frame, text="Ekip seç", bg=CANVAS, fg=INK, font=(FONT, 12, "bold")).pack(side="left", padx=(0, 8))
        team_names = [t["name"] for t in teams]
        selected_team = tk.StringVar(value=team_names[0] if team_names else "")
        team_combo = ttk.Combobox(
            filter_frame, textvariable=selected_team, values=team_names,
            state="readonly", font=(FONT, 12), width=28,
        )
        team_combo.pack(side="left", ipady=4)

        listing = tk.Frame(self.content, bg=CANVAS)
        listing.pack(fill="both", expand=True, padx=20, pady=(4, 16))

        def paint(*_args):
            for child in listing.winfo_children():
                child.destroy()
            chosen = selected_team.get()
            all_codes = [
                item for item in self.store.request_codes()
                if item["team_name"] == chosen or item["all_teams"]
            ]
            if not all_codes:
                tk.Label(listing, text="Bu ekipte henüz talep kodu yok.", bg=CANVAS, fg=MUTED, font=(FONT, 13)).pack(anchor="w", pady=12)
                return
            head = tk.Frame(listing, bg=CANVAS)
            head.pack(fill="x", pady=(8, 4))
            for text, width in (("Talep kodu", 24), ("Başlık", 18), ("Tür", 8), ("İş kodu", 12)):
                tk.Label(head, text=text, bg=CANVAS, fg=MUTED, font=(FONT, 11, "bold"), width=width, anchor="w").pack(side="left")
            for item in all_codes:
                row = tk.Frame(listing, bg=CARD, highlightbackground=LINE, highlightthickness=1)
                row.pack(fill="x", pady=2)
                shown_title = item["title"] if "title" in item.keys() and item["title"] else "—"
                shown_scope = item["scope"] if "scope" in item.keys() and item["scope"] else "İç"
                shown_code = f"{item['code']}  ·  Hepsi → {item['team_name']}" if item["all_teams"] else item["code"]
                for value, width in ((shown_code, 24), (shown_title, 18), (shown_scope, 8), (item["work_code"], 12)):
                    tk.Label(row, text=value, bg=CARD, fg=INK, font=(FONT, 12), width=width, anchor="w").pack(side="left", padx=(8, 0), pady=6)
                Click(row, "✎", lambda current=item: self._edit_code(current, paint), CARD, ACCENT, HOVER, padx=10, pady=4).pack(side="right", padx=6)

        team_combo.bind("<<ComboboxSelected>>", paint)
        paint()

    def _edit_code(self, item, refresh):
        dialog = tk.Toplevel(self)
        dialog.title("Talep kodunu düzenle")
        dialog.configure(bg=CANVAS)
        dialog.transient(self)
        dialog.resizable(False, False)
        card = tk.Frame(dialog, bg=CARD, padx=16, pady=14, highlightbackground=LINE, highlightthickness=1)
        card.pack(padx=12, pady=12)
        tk.Label(card, text="Talep kodunu düzenle", bg=CARD, fg=INK, font=(FONT, 16, "bold")).pack(anchor="w")
        code = tk.StringVar(value=item["code"])
        heading = tk.StringVar(value=item["title"] if "title" in item.keys() and item["title"] else "")
        work = tk.StringVar(value=item["work_code"])
        teams = list(self.store.teams())
        shared = bool(item["all_teams"]) if "all_teams" in item.keys() else False
        current = next((team for team in teams if team["id"] == item["team_id"]), None)
        membership = tk.StringVar(value="Hepsi" if shared else (self._team_label(current) if current else ""))
        for label, var in (("Talep kodu", code), ("Başlık", heading)):
            tk.Label(card, text=label, bg=CARD, fg=MUTED, font=(FONT, 11, "bold")).pack(anchor="w", pady=(8, 2))
            tk.Entry(card, textvariable=var, relief="flat", highlightthickness=1, highlightbackground=LINE, font=(FONT, 13), width=36).pack(fill="x", ipady=5)
        tk.Label(card, text="Ekip", bg=CARD, fg=MUTED, font=(FONT, 11, "bold")).pack(anchor="w", pady=(8, 2))
        ttk.Combobox(
            card, textvariable=membership, values=["Hepsi"] + [self._team_label(team) for team in teams],
            state="readonly", font=(FONT, 12), width=34,
        ).pack(fill="x", ipady=4)
        edit_target = tk.Frame(card, bg=CARD)
        tk.Label(edit_target, text="Hedef ekip", bg=CARD, fg=MUTED, font=(FONT, 11, "bold")).pack(anchor="w", pady=(8, 2))
        target_var = tk.StringVar(value=self._team_label(current) if current else "")
        ttk.Combobox(
            edit_target, textvariable=target_var, values=[self._team_label(team) for team in teams],
            state="readonly", font=(FONT, 12), width=34,
        ).pack(fill="x", ipady=4)
        tk.Label(edit_target, text="Kod bu ekibe düşer. Her ekip bu kodu kendi listesinde görür.", bg=CARD, fg=MUTED, font=(FONT, 11)).pack(anchor="w", pady=(4, 0))
        edit_scope_label = tk.Label(card, text="Tür", bg=CARD, fg=MUTED, font=(FONT, 11, "bold"))
        edit_scope_label.pack(anchor="w", pady=(8, 2))

        def paint_share(*_args):
            if membership.get() == "Hepsi":
                edit_target.pack(fill="x", before=edit_scope_label)
            else:
                edit_target.pack_forget()

        membership.trace_add("write", paint_share)
        paint_share()
        edit_scope = tk.StringVar(value=item["scope"] if "scope" in item.keys() and item["scope"] else "İç")
        edit_scope_frame = tk.Frame(card, bg=CARD)
        edit_scope_frame.pack(anchor="w")
        tk.Radiobutton(edit_scope_frame, text="İç", variable=edit_scope, value="İç", bg=CARD, fg=INK, font=(FONT, 12), selectcolor=CARD, activebackground=CARD).pack(side="left", padx=(0, 16))
        tk.Radiobutton(edit_scope_frame, text="Dış", variable=edit_scope, value="Dış", bg=CARD, fg=INK, font=(FONT, 12), selectcolor=CARD, activebackground=CARD).pack(side="left")
        tk.Label(card, text="İş kodu", bg=CARD, fg=MUTED, font=(FONT, 11, "bold")).pack(anchor="w", pady=(8, 2))
        tk.Entry(card, textvariable=work, relief="flat", highlightthickness=1, highlightbackground=LINE, font=(FONT, 13), width=36).pack(fill="x", ipady=5)

        def save():
            shared_now = membership.get() == "Hepsi"
            picked = target_var.get() if shared_now else membership.get()
            team = self._team_from_label(picked, self.store.teams())
            if team is None:
                messagebox.showwarning(APP_NAME, "Hedef ekip seç." if shared_now else "Ekip seç.", parent=dialog)
                return
            try:
                self.store.update_request_code(item["id"], code.get(), team["id"], work.get(), heading.get(), edit_scope.get(), shared_now)
            except ValueError as exc:
                messagebox.showwarning(APP_NAME, str(exc), parent=dialog)
                return
            dialog.destroy()
            refresh()

        Click(card, "Kaydet", save, ACCENT, "white", ACCENT_HOVER).pack(anchor="e", pady=(14, 0))
        dialog.bind("<Escape>", lambda _e: dialog.destroy())
        dialog.grab_set()

    def _build_logs(self):
        if not self.user["is_admin"]:
            self._show_page("surec")
            return
        head = tk.Frame(self.content, bg=CANVAS)
        head.pack(fill="x", padx=20, pady=(16, 2))
        tk.Label(head, text="Giriş kayıtları", bg=CANVAS, fg=INK, font=(FONT, 22, "bold")).pack(side="left")
        Click(head, "Yenile", lambda: self._fill_logs(), ACCENT, "white", ACCENT_HOVER).pack(side="right")
        tk.Label(
            self.content,
            text="Giriş, çıkış ve başarısız deneme. Cihaz adı, sistem ve bu bilgisayarın IP adresi satıra yazılır.",
            bg=CANVAS, fg=MUTED, font=(FONT, 12), wraplength=760, justify="left",
        ).pack(anchor="w", padx=20)
        stats = tk.Frame(self.content, bg=CANVAS)
        stats.pack(fill="x", padx=20, pady=(12, 0))
        self._log_today = tk.StringVar(value="0")
        self._log_failed = tk.StringVar(value="0")
        self._log_devices = tk.StringVar(value="0")
        for title, var in (("Bugünkü giriş", self._log_today), ("Bugünkü başarısız", self._log_failed), ("Farklı cihaz", self._log_devices)):
            card = tk.Frame(stats, bg=CARD, highlightbackground=LINE, highlightthickness=1, padx=14, pady=10)
            card.pack(side="left", fill="x", expand=True, padx=(0, 8))
            tk.Label(card, text=title, bg=CARD, fg=MUTED, font=(FONT, 11, "bold")).pack(anchor="w")
            tk.Label(card, textvariable=var, bg=CARD, fg=INK, font=(FONT, 18, "bold")).pack(anchor="w", pady=(2, 0))
        filters = tk.Frame(self.content, bg=CANVAS)
        filters.pack(fill="x", padx=20, pady=(12, 0))
        self._log_event = tk.StringVar(value="Tümü")
        tk.Label(filters, text="Olay", bg=CANVAS, fg=MUTED, font=(FONT, 11, "bold")).pack(side="left", padx=(0, 6))
        ttk.Combobox(
            filters, textvariable=self._log_event, values=("Tümü",) + LOG_EVENTS,
            state="readonly", font=(FONT, 12), width=16,
        ).pack(side="left", ipady=3)
        self._log_query = tk.StringVar()
        tk.Label(filters, text="Ara", bg=CANVAS, fg=MUTED, font=(FONT, 11, "bold")).pack(side="left", padx=(12, 6))
        tk.Entry(
            filters, textvariable=self._log_query, relief="flat", highlightthickness=1, highlightbackground=LINE,
            font=(FONT, 12),
        ).pack(side="left", fill="x", expand=True, padx=(8, 0), ipady=5)
        self._log_detail = tk.Frame(self.content, bg=CARD, highlightbackground=LINE, highlightthickness=1, padx=14, pady=10)
        self._log_detail.pack(side="bottom", fill="x", padx=20, pady=(8, 16))
        table = tk.Frame(self.content, bg=CARD, highlightbackground=LINE, highlightthickness=1)
        table.pack(fill="both", expand=True, padx=20, pady=(10, 0))
        columns = (("when", "Zaman", 150), ("name", "Kullanıcı", 150), ("email", "E-posta", 210), ("event", "Olay", 100), ("device", "Cihaz", 180), ("ip", "IP", 120))
        self._log_tree = ttk.Treeview(table, columns=tuple(key for key, _text, _width in columns), show="headings")
        for key, text, width in columns:
            self._log_tree.heading(key, text=text)
            self._log_tree.column(key, width=width, anchor="w")
        self._log_tree.tag_configure("Giriş", foreground=OK)
        self._log_tree.tag_configure("Çıkış", foreground=MUTED)
        self._log_tree.tag_configure("Başarısız", foreground=DANGER)
        scroll = ttk.Scrollbar(table, orient="vertical", command=self._log_tree.yview)
        self._log_tree.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        self._log_tree.pack(side="left", fill="both", expand=True)
        self._log_tree.bind("<<TreeviewSelect>>", lambda _e: self._show_log_detail())
        self._log_rows = {}
        self._log_event.trace_add("write", lambda *_a: self._fill_logs())
        self._log_query.trace_add("write", lambda *_a: self._fill_logs())
        self._fill_logs()

    def _fill_logs(self):
        if not hasattr(self, "_log_tree") or not self._log_tree.winfo_exists():
            return
        rows = list(self.store.login_logs())
        today = datetime.now().strftime("%Y-%m-%d")
        self._log_today.set(str(sum(1 for row in rows if row["event"] == "Giriş" and str(row["created_at"]).startswith(today))))
        self._log_failed.set(str(sum(1 for row in rows if row["event"] == "Başarısız" and str(row["created_at"]).startswith(today))))
        self._log_devices.set(str(len({row["device"] for row in rows if row["device"]})))
        wanted = self._log_event.get()
        query = " ".join(self._log_query.get().casefold().split())
        for item in self._log_tree.get_children():
            self._log_tree.delete(item)
        self._log_rows = {}
        for row in rows:
            if wanted != "Tümü" and row["event"] != wanted:
                continue
            blob = " ".join(
                str(row[key] or "") for key in ("name", "email", "event", "device", "os_name", "ip")
            ).casefold()
            if query and query not in blob:
                continue
            iid = str(row["id"])
            self._log_tree.insert(
                "", "end", iid=iid, tags=(row["event"],),
                values=(row["created_at"], row["name"] or "—", row["email"] or "—", row["event"], row["device"] or "—", row["ip"] or "—"),
            )
            self._log_rows[row["id"]] = row
        self._show_log_detail()

    def _show_log_detail(self):
        if not hasattr(self, "_log_detail") or not self._log_detail.winfo_exists():
            return
        for child in self._log_detail.winfo_children():
            child.destroy()
        selected = self._log_tree.selection() if self._log_tree.winfo_exists() else ()
        row = self._log_rows.get(int(selected[0])) if selected else None
        if row is None:
            tk.Label(self._log_detail, text="Bir satır seçince cihaz, sistem ve IP burada durur.", bg=CARD, fg=MUTED, font=(FONT, 12)).pack(anchor="w")
            return
        who = row["name"] or "Tanınmayan deneme"
        tk.Label(self._log_detail, text=f"{who}  ·  {row['email'] or '—'}", bg=CARD, fg=INK, font=(FONT, 14, "bold")).pack(anchor="w")
        tone = {"Giriş": OK, "Çıkış": MUTED, "Başarısız": DANGER}.get(row["event"], INK)
        tk.Label(self._log_detail, text=f"{row['event']}  ·  {row['created_at']}", bg=CARD, fg=tone, font=(FONT, 12, "bold")).pack(anchor="w", pady=(2, 0))
        device = "  ·  ".join(part for part in (row["device"], row["os_name"], row["ip"] or "IP yok") if part)
        tk.Label(self._log_detail, text=device or "Cihaz bilgisi yok", bg=CARD, fg=INK, font=(FONT, 12)).pack(anchor="w", pady=(2, 0))

    def _build_users(self):
        if not self.user["is_admin"]:
            self._show_page("talepler")
            return
        self._page_head("Kullanıcılar", "Hesap bir rolle açılır. Yönetici yalnızca furkan.hamarat01@gmail.com hesabıdır. Diğer hesaplar Kullanıcı, Sorumlu veya İzleyici olur.")
        teams = list(self.store.teams())
        card = self._form_card()
        if not teams:
            tk.Label(card, text="Henüz ekip yok.", bg=CARD, fg=INK, font=(FONT, 13)).pack(anchor="w")
            Click(card, "Ekiplere git", lambda: self._show_page("ekipler"), ACCENT, "white", ACCENT_HOVER).pack(anchor="e", pady=(12, 0))
        else:
            person = self._line(card, "Ad soyad")
            username = self._line(card, "Kullanıcı adı")
            mail = self._line(card, "E-posta")
            secret_password = self._line(card, "Parola", secret=True)
            tk.Label(card, text="Ekip", bg=CARD, fg=MUTED, font=(FONT, 11, "bold")).pack(anchor="w", pady=(8, 2))
            membership = tk.StringVar(value=self._team_label(teams[0]))
            ttk.Combobox(
                card, textvariable=membership, values=[self._team_label(team) for team in teams],
                state="readonly", font=(FONT, 12),
            ).pack(fill="x", ipady=4)
            tk.Label(card, text="Rol", bg=CARD, fg=MUTED, font=(FONT, 11, "bold")).pack(anchor="w", pady=(8, 2))
            role_var = tk.StringVar(value=ROLES[0])
            ttk.Combobox(
                card, textvariable=role_var, values=list(ROLES),
                state="readonly", font=(FONT, 12),
            ).pack(fill="x", ipady=4)

            def add_user():
                team = self._team_from_label(membership.get(), self.store.teams())
                if team is None:
                    messagebox.showwarning(APP_NAME, "Ekip seç.", parent=self)
                    return
                if role_var.get() not in ROLES:
                    messagebox.showwarning(APP_NAME, "Rol seç.", parent=self)
                    return
                secret = new_totp_secret()
                try:
                    user_id = self.store.add_user(person.get(), mail.get(), secret_password.get(), secret, False, username.get(), role=role_var.get())
                    self.store.add_member(user_id, team["id"])
                except ValueError as exc:
                    messagebox.showwarning(APP_NAME, str(exc), parent=self)
                    return
                created = self.store.user_by_email(mail.get())

                def finish(user=created, key=secret):
                    self._show_page("kullanicilar")
                    self._handoff_secret(user, key)

                self.after_idle(finish)

            Click(card, "Kullanıcı ekle", add_user, ACCENT, "white", ACCENT_HOVER).pack(anchor="e", pady=(12, 0))
        listing = tk.Frame(self.content, bg=CANVAS)
        listing.pack(fill="both", expand=True, padx=20, pady=(4, 16))
        head = tk.Frame(listing, bg=CANVAS)
        head.pack(fill="x", pady=(0, 4))
        for text, width in (("Ad", 16), ("Kullanıcı adı", 14), ("E-posta", 22), ("Ekip", 14), ("Rol", 12)):
            tk.Label(head, text=text, bg=CANVAS, fg=MUTED, font=(FONT, 11, "bold"), width=width, anchor="w").pack(side="left")
        for user in self.store.users():
            row = tk.Frame(listing, bg=CARD, highlightbackground=LINE, highlightthickness=1)
            row.pack(fill="x", pady=2)
            names = ", ".join(self._team_label(team) for team in self.store.user_teams(user["id"])) or "—"
            handle = user["username"] if "username" in user.keys() and user["username"] else "—"
            role = user["role"] if "role" in user.keys() and user["role"] else ("Yönetici" if user["is_admin"] else "Kullanıcı")
            for value, width in ((user["name"], 16), (handle, 14), (user["email"], 22), (names, 14), (role, 12)):
                fg = ACCENT if value == ADMIN_ROLE else INK
                tk.Label(row, text=value, bg=CARD, fg=fg, font=(FONT, 12, "bold") if value == ADMIN_ROLE else (FONT, 12), width=width, anchor="w").pack(side="left", padx=(8, 0), pady=6)
            Click(row, "Ayarlar", lambda item=user: self._edit_user(item), CARD, ACCENT, HOVER, padx=10, pady=4).pack(side="right", padx=6)

    def _edit_user(self, user):
        dialog = tk.Toplevel(self)
        dialog.title("Kullanıcı ayarları")
        dialog.configure(bg=CANVAS)
        dialog.transient(self)
        dialog.resizable(False, False)
        card = tk.Frame(dialog, bg=CARD, padx=16, pady=14, highlightbackground=LINE, highlightthickness=1)
        card.pack(padx=12, pady=12)
        tk.Label(card, text="Kullanıcı ayarları", bg=CARD, fg=INK, font=(FONT, 16, "bold")).pack(anchor="w")
        tk.Label(
            card, text="Geçici parola yazarsan kullanıcı bir kez onunla girer, sonra kendi parolasını belirler.",
            bg=CARD, fg=MUTED, font=(FONT, 12), wraplength=420, justify="left",
        ).pack(anchor="w", pady=(4, 4))
        name = tk.StringVar(value=user["name"])
        username = tk.StringVar(value=user["username"] if "username" in user.keys() and user["username"] else "")
        email = tk.StringVar(value=user["email"])
        temporary = tk.StringVar()
        for label, var, secret in (
            ("Ad soyad", name, False),
            ("Kullanıcı adı", username, False),
            ("E-posta", email, False),
            ("Geçici parola", temporary, True),
        ):
            tk.Label(card, text=label, bg=CARD, fg=MUTED, font=(FONT, 11, "bold")).pack(anchor="w", pady=(8, 2))
            entry = tk.Entry(card, textvariable=var, show="•" if secret else "", relief="flat", highlightthickness=1, highlightbackground=LINE, font=(FONT, 13), width=40)
            if label == "E-posta" and clean_email(user["email"]) == ADMIN_EMAIL:
                entry.configure(state="readonly", readonlybackground=CARD)
            entry.pack(fill="x", ipady=5)
        teams = list(self.store.teams())
        current_teams = self.store.user_teams(user["id"])
        membership = tk.StringVar(value=self._team_label(current_teams[0]) if current_teams else (self._team_label(teams[0]) if teams else ""))
        tk.Label(card, text="Ekip", bg=CARD, fg=MUTED, font=(FONT, 11, "bold")).pack(anchor="w", pady=(8, 2))
        if teams:
            ttk.Combobox(
                card, textvariable=membership, values=[self._team_label(team) for team in teams],
                state="readonly", font=(FONT, 12), width=38,
            ).pack(fill="x", ipady=4)
        current_role = user["role"] if "role" in user.keys() and user["role"] else ("Yönetici" if user["is_admin"] else "Kullanıcı")
        locked_admin = clean_email(user["email"]) == ADMIN_EMAIL
        tk.Label(card, text="Rol", bg=CARD, fg=MUTED, font=(FONT, 11, "bold")).pack(anchor="w", pady=(8, 2))
        role_var = tk.StringVar(value=ADMIN_ROLE if locked_admin else (current_role if current_role in ROLES else ROLES[0]))
        if locked_admin:
            tk.Label(card, text="Yönetici", bg=CARD, fg=ACCENT, font=(FONT, 13, "bold")).pack(anchor="w")
        else:
            ttk.Combobox(
                card, textvariable=role_var, values=list(ROLES),
                state="readonly", font=(FONT, 12), width=38,
            ).pack(fill="x", ipady=4)

        def save():
            team = self._team_from_label(membership.get(), self.store.teams()) if teams else None
            try:
                self.store.update_user(user["id"], name.get(), email.get(), username.get(), team["id"] if team else None, role_var.get())
                if temporary.get().strip():
                    self.store.set_password(user["id"], temporary.get(), True)
            except ValueError as exc:
                messagebox.showwarning(APP_NAME, str(exc), parent=dialog)
                return
            dialog.destroy()
            self._show_page("kullanicilar")

        Click(card, "Kaydet", save, ACCENT, "white", ACCENT_HOVER).pack(anchor="e", pady=(14, 0))
        dialog.bind("<Escape>", lambda _e: dialog.destroy())
        dialog.grab_set()

    def _handoff_secret(self, user, secret):
        folder = app_dir() / "anahtarlar"
        folder.mkdir(parents=True, exist_ok=True)
        note = folder / f"{user['email']}.txt"
        note.write_text(
            f"{user['name']}\n{user['email']}\n{secret}\n\n"
            "Kimlik uygulamasında Hesap ekle → Anahtarı gir.\n",
            encoding="utf-8",
        )
        try:
            os.chmod(note, 0o600)
        except OSError:
            pass
        dialog = tk.Toplevel(self)
        dialog.title("Kurulum anahtarı")
        dialog.configure(bg=CANVAS)
        dialog.transient(self)
        dialog.resizable(False, False)
        card = tk.Frame(dialog, bg=CARD, padx=18, pady=16, highlightbackground=LINE, highlightthickness=1)
        card.pack(padx=12, pady=12)
        tk.Label(card, text=user["name"], bg=CARD, fg=INK, font=(FONT, 16, "bold")).pack(anchor="w")
        tk.Label(
            card,
            text="Bu anahtar telefona bir kez eklenir. Hesap ekle, Anahtarı gir, yapıştır. Sonra kullanıcı e-posta, parola ve 6 hane ile girer.",
            bg=CARD, fg=MUTED, font=(FONT, 12), wraplength=420, justify="left",
        ).pack(anchor="w", pady=(4, 10))
        shown = tk.Entry(card, relief="flat", highlightthickness=1, highlightbackground=LINE, font=(FONT, 14), width=36, readonlybackground=CARD)
        shown.insert(0, secret)
        shown.configure(state="readonly")
        shown.pack(fill="x", ipady=6)
        actions = tk.Frame(card, bg=CARD)
        actions.pack(fill="x", pady=(12, 0))
        Click(actions, "Kopyala", lambda: self._copy(secret), SOFT, ACCENT, HOVER).pack(side="left")
        Click(actions, "Kapat", dialog.destroy, ACCENT, "white", ACCENT_HOVER).pack(side="right")
        dialog.bind("<Escape>", lambda _e: dialog.destroy())
        dialog.grab_set()
        try:
            shown.selection_range(0, "end")
        except tk.TclError:
            pass


def _self_check():
    assert "287082" in totp_codes("GEZDGNBVGY3TQOJQGEZDGNBVGY3TQOJQ", 59, window=0)
    assert code_from_topic("Satın alma") == "SA"
    assert code_from_topic("Muhasebe") == "MUHA"
    assert code_from_topic("Bilgi teknolojileri") == "BT"
    stored = hash_password("gizli-parola")
    assert check_password("gizli-parola", stored)
    assert not check_password("yanlis", stored)
    print("ok")


def main():
    if "--check" in sys.argv:
        _self_check()
        return
    TalepApp().mainloop()


if __name__ == "__main__":
    main()
