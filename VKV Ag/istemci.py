# -*- coding: utf-8 -*-
"""VKV Ağ istemcisi. Her bilgisayarda ayrı açılır ve sunucuya bağlanır.

İş Merkezi ve Talep kayıtlarına yazmaz. Talep, mesaj ve dosya sunucuda durur.
Bu bilgisayarın adresi ve oturumu yalnız bu kullanıcı hesabında saklanır.
"""
import json
import sys
import threading
import tkinter as tk
import urllib.error
import urllib.parse
import urllib.request
import uuid
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from ortak import (
    APP, FONT, FILE_EXT, MAX_FILE, PUBLIC_SERVER, ROLES, STATUSES,
    api_base, device_snapshot, file_size, istemci_dir, safe_filename,
)

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


class ApiError(Exception):
    pass


class Api:
    def __init__(self):
        self.base = ""
        self.token = ""
        self.client_id = ""

    def set_base(self, text):
        self.base = api_base(text)

    def durum(self):
        return self._request("/v1/durum", auth=False)

    def kurulum(self, key, name, username, password):
        return self._request("/v1/kurulum", {"anahtar": key, "ad": name, "kullanici": username, "parola": password}, auth=False)["kullanici"]

    def giris(self, username, password):
        user = self._request("/v1/giris", {"kullanici": username, "parola": password}, auth=False)["kullanici"]
        self.token = user["token"]
        return user

    def cikis(self):
        try:
            self._request("/v1/cikis", {})
        finally:
            self.token = ""

    def beat(self, device):
        self._request("/v1/nabiz", device)

    def users(self):
        return self._request("/v1/kullanicilar")["kullanicilar"]

    def add_user(self, name, username, password, role):
        return self._request("/v1/kullanici", {"ad": name, "kullanici": username, "parola": password, "rol": role})["kullanici"]

    def computers(self):
        return self._request("/v1/bilgisayarlar")["bilgisayarlar"]

    def distribution(self):
        return self._request("/v1/dagilim")["dagilim"]

    def requests(self):
        return self._request("/v1/talepler")["talepler"]

    def create_request(self, title, note, assignee_id):
        return self._request("/v1/talep", {"baslik": title, "not": note, "kime": assignee_id})["talep"]

    def set_status(self, request_id, status):
        return self._request("/v1/talep/durum", {"id": request_id, "durum": status})["talep"]

    def files(self):
        return self._request("/v1/dosyalar")["dosyalar"]

    def upload(self, path, note=""):
        source = Path(path)
        clean = safe_filename(source.name)
        payload = source.read_bytes()
        if not payload:
            raise ApiError("Dosya boş.")
        if len(payload) > MAX_FILE:
            raise ApiError("Dosya 20 MB sınırını aşıyor.")
        return self._request("/v1/dosya", raw=payload, headers={
            "X-Dosya-Ad": urllib.parse.quote(clean),
            "X-Dosya-Not": urllib.parse.quote(" ".join(str(note or "").split())[:200]),
        })["dosya"]

    def download(self, file_id):
        return self._request(f"/v1/dosya?id={int(file_id)}")

    def messages(self):
        return self._request("/v1/mesajlar")["mesajlar"]

    def send_message(self, text):
        return self._request("/v1/mesaj", {"metin": text})["mesaj"]

    def _request(self, path, payload=None, raw=None, headers=None, auth=True):
        if not self.base:
            raise ApiError("Sunucu adresi yok.")
        data = raw if raw is not None else (None if payload is None else json.dumps(payload, ensure_ascii=False).encode("utf-8"))
        request = urllib.request.Request(self.base + path, data=data, method="GET" if data is None else "POST")
        if payload is not None and raw is None:
            request.add_header("Content-Type", "application/json; charset=utf-8")
        if auth and self.token:
            request.add_header("Authorization", f"Bearer {self.token}")
        if self.client_id:
            request.add_header("X-Istemci", self.client_id)
        for key, value in (headers or {}).items():
            request.add_header(key, value)
        try:
            with urllib.request.urlopen(request, timeout=8) as response:
                body = response.read()
                if "json" in response.headers.get("Content-Type", ""):
                    parsed = json.loads(body.decode("utf-8"))
                    if not parsed.get("ok", True):
                        raise ApiError(parsed.get("error") or "İstek olmadı.")
                    return parsed
                return body
        except urllib.error.HTTPError as exc:
            try:
                parsed = json.loads(exc.read().decode("utf-8"))
            except (UnicodeError, json.JSONDecodeError):
                parsed = {}
            if exc.code == 401:
                self.token = ""
            raise ApiError(parsed.get("error") or "İstek olmadı.") from exc
        except urllib.error.URLError as exc:
            raise ApiError("Sunucuya ulaşılamıyor.") from exc


class Click(tk.Frame):
    def __init__(self, master, text, command, bg, fg, hover=None, padx=14, pady=8, anchor="center"):
        super().__init__(master, bg=bg, cursor="hand2")
        self.command = command
        self._bg = bg
        self._fg = fg
        self._hover = hover or bg
        self.label = tk.Label(self, text=text, bg=bg, fg=fg, font=(FONT, 12, "bold"), padx=padx, pady=pady, cursor="hand2", anchor=anchor)
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


class AgApp(tk.Tk):
    def __init__(self, start=True):
        super().__init__()
        self.title(APP)
        self.configure(bg=CANVAS)
        self.minsize(980, 640)
        self.geometry("1100x720")
        self.api = Api()
        self.config_data = self._load_config()
        self.api.client_id = self.config_data["istemci"]
        self.api.token = self.config_data.get("token") or ""
        self.user = None
        self.page = ""
        self._generation = 0
        self._busy = False
        self._nav = {}
        self.shell = tk.Frame(self, bg=CANVAS)
        self.shell.pack(fill="both", expand=True)
        self.protocol("WM_DELETE_WINDOW", self._close)
        if start:
            self.after(80, self._boot)

    def _config_path(self):
        return istemci_dir() / "ayar.json"

    def _load_config(self):
        path = self._config_path()
        data = {}
        if path.is_file():
            try:
                data = json.loads(path.read_text(encoding="utf-8"))
            except (OSError, json.JSONDecodeError):
                data = {}
        if not data.get("istemci"):
            data["istemci"] = uuid.uuid4().hex
            self._write_config(data)
        return data

    def _write_config(self, data):
        path = self._config_path()
        path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        try:
            path.chmod(0o600)
        except OSError:
            pass

    def _save_config(self):
        self.config_data["token"] = self.api.token
        self.config_data["istemci"] = self.api.client_id
        self._write_config(self.config_data)

    def _close(self):
        self._generation += 1
        self.destroy()

    def _clear(self):
        for child in self.shell.winfo_children():
            child.destroy()

    def _boot(self):
        saved = (self.config_data.get("sunucu") or "").strip().rstrip("/")
        if saved in ("", "127.0.0.1:8780", "http://127.0.0.1:8780"):
            saved = PUBLIC_SERVER
        self._connect_screen(saved, True)

    def _connect_screen(self, address, auto):
        self._clear()
        card = self._card(
            "Sunucuya bağlan",
            "Bu program her bilgisayarda ayrıdır. Yedek sunucudaki adresi yaz. İş Merkezi ve Talep kayıtları burada durmaz.",
        )
        tk.Label(card, text="Sunucu adresi", bg=CARD, fg=MUTED, font=(FONT, 11, "bold")).pack(anchor="w", pady=(8, 2))
        self.address = tk.StringVar(value=address)
        tk.Entry(card, textvariable=self.address, relief="flat", highlightthickness=1, highlightbackground=LINE, font=(FONT, 13), width=62).pack(fill="x", ipady=6)
        Click(card, "Bağlan", self._connect, ACCENT, "white", ACCENT_HOVER).pack(anchor="e", pady=(16, 0))
        if auto:
            self.after(60, self._connect)

    def _connect(self):
        try:
            self.api.set_base(self.address.get())
        except ValueError as exc:
            messagebox.showwarning(APP, str(exc), parent=self)
            return
        self.config_data["sunucu"] = self.address.get().strip()
        self._save_config()

        def work():
            try:
                durum = self.api.durum()
            except ApiError as exc:
                self.after(0, lambda: messagebox.showwarning(APP, str(exc), parent=self))
                return
            self.after(0, lambda: self._after_connect(durum))

        threading.Thread(target=work, daemon=True).start()

    def _after_connect(self, durum):
        if durum.get("kurulum"):
            self._setup_screen()
            return
        if self.api.token:
            def work():
                try:
                    self.api.beat(device_snapshot())
                    user = {"ad": self.config_data.get("ad") or "", "kullanici": self.config_data.get("kullanici") or "", "rol": self.config_data.get("rol") or "Kullanıcı", "id": self.config_data.get("id") or 0, "token": self.api.token}
                except ApiError:
                    self.api.token = ""
                    self._save_config()
                    self.after(0, self._login_screen)
                    return
                self.after(0, lambda: self._enter(user))

            threading.Thread(target=work, daemon=True).start()
            return
        self._login_screen()

    def _card(self, title, note):
        card = tk.Frame(self.shell, bg=CARD, highlightbackground=LINE, highlightthickness=1, padx=22, pady=18)
        card.place(relx=0.5, rely=0.5, anchor="center")
        tk.Label(card, text=title, bg=CARD, fg=INK, font=(FONT, 20, "bold")).pack(anchor="w")
        tk.Label(card, text=note, bg=CARD, fg=MUTED, font=(FONT, 12), wraplength=440, justify="left").pack(anchor="w", pady=(4, 12))
        return card

    def _field(self, parent, label, secret=False):
        tk.Label(parent, text=label, bg=CARD, fg=MUTED, font=(FONT, 11, "bold")).pack(anchor="w", pady=(8, 2))
        var = tk.StringVar()
        tk.Entry(parent, textvariable=var, relief="flat", highlightthickness=1, highlightbackground=LINE, font=(FONT, 13), show="•" if secret else "").pack(fill="x", ipady=6)
        return var

    def _setup_screen(self):
        self._clear()
        card = self._card("İlk yönetici", "Sunucu penceresindeki kurulum anahtarını yaz. Bu hesap diğer kişileri ekler.")
        key = self._field(card, "Kurulum anahtarı")
        name = self._field(card, "Ad soyad")
        username = self._field(card, "Kullanıcı adı")
        password = self._field(card, "Parola", secret=True)

        def save():
            def work():
                try:
                    self.api.kurulum(key.get(), name.get(), username.get(), password.get())
                    user = self.api.giris(username.get(), password.get())
                except ApiError as exc:
                    self.after(0, lambda: messagebox.showwarning(APP, str(exc), parent=self))
                    return
                self.after(0, lambda: self._enter(user))

            threading.Thread(target=work, daemon=True).start()

        Click(card, "Yöneticiyi kur", save, ACCENT, "white", ACCENT_HOVER).pack(anchor="e", pady=(16, 0))

    def _login_screen(self):
        self._clear()
        card = self._card("Giriş", "Bu bilgisayarın kendi oturumu. Sunucudaki kullanıcı adın ve parolan ile gir.")
        username = self._field(card, "Kullanıcı adı")
        password = self._field(card, "Parola", secret=True)
        if self.config_data.get("kullanici"):
            username.set(self.config_data["kullanici"])

        def enter():
            def work():
                try:
                    user = self.api.giris(username.get(), password.get())
                except ApiError as exc:
                    self.after(0, lambda: messagebox.showwarning(APP, str(exc), parent=self))
                    return
                self.after(0, lambda: self._enter(user))

            threading.Thread(target=work, daemon=True).start()

        Click(card, "Giriş yap", enter, ACCENT, "white", ACCENT_HOVER).pack(anchor="e", pady=(16, 0))
        Click(card, "Sunucu adresini değiştir", lambda: self._connect_screen(self.config_data.get("sunucu") or "", False), CARD, ACCENT, HOVER).pack(anchor="e", pady=(8, 0))

    def _enter(self, user):
        self._generation += 1
        self.user = user
        self.config_data["ad"] = user.get("ad") or ""
        self.config_data["kullanici"] = user.get("kullanici") or ""
        self.config_data["rol"] = user.get("rol") or ""
        self.config_data["id"] = user.get("id") or 0
        self._save_config()
        self._clear()
        sidebar = tk.Frame(self.shell, bg=NAV, width=220)
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)
        tk.Label(sidebar, text=APP, bg=NAV, fg="white", font=(FONT, 16, "bold")).pack(anchor="w", padx=16, pady=(18, 2))
        tk.Label(sidebar, text=user.get("ad") or "", bg=NAV, fg="white", font=(FONT, 12)).pack(anchor="w", padx=16)
        tk.Label(sidebar, text=user.get("rol") or "", bg=NAV, fg="#8FDBA8", font=(FONT, 11, "bold")).pack(anchor="w", padx=16, pady=(2, 14))
        self._nav = {}
        self._nav_item(sidebar, "dagilim", "Talep dağılımı")
        self._nav_item(sidebar, "talepler", "Talepler")
        self._nav_item(sidebar, "dosyalar", "Dosyalar")
        self._nav_item(sidebar, "mesajlar", "Mesajlar")
        self._nav_item(sidebar, "bilgisayarlar", "Bilgisayarlar")
        foot = tk.Frame(sidebar, bg=NAV)
        foot.pack(side="bottom", fill="x", pady=12)
        device = device_snapshot()
        self.link = tk.Label(foot, text="Sunucuya bağlanıyor", bg=NAV, fg=NAV_TEXT, font=(FONT, 11), wraplength=180, justify="left")
        self.link.pack(anchor="w", padx=16, pady=(0, 4))
        tk.Label(foot, text=device["cihaz"], bg=NAV, fg=NAV_TEXT, font=(FONT, 11), wraplength=180, justify="left").pack(anchor="w", padx=16, pady=(0, 8))
        Click(foot, "Çıkış", self._logout, NAV, NAV_TEXT, "#2E3A54", anchor="w", padx=16).pack(fill="x", padx=12)
        self.content = tk.Frame(self.shell, bg=CANVAS)
        self.content.pack(side="left", fill="both", expand=True)
        self._show("dagilim")
        generation = self._generation
        self.after(4000, lambda: self._loop(generation))

    def _nav_item(self, parent, key, text):
        button = Click(parent, text, lambda: self._show(key), NAV, NAV_TEXT, "#2E3A54", anchor="w", padx=16, pady=10)
        button.pack(fill="x", padx=12, pady=2)
        self._nav[key] = button

    def _paint_nav(self):
        for key, button in self._nav.items():
            if key == self.page:
                button.set_palette("#2E3A54", "white", "#3A4866")
            else:
                button.set_palette(NAV, NAV_TEXT, "#2E3A54")

    def _logout(self):
        self._generation += 1
        token = self.api.token
        self.api.token = ""
        self.user = None
        self._save_config()
        if token:
            threading.Thread(target=lambda: self._drop(token), daemon=True).start()
        self._login_screen()

    def _drop(self, token):
        previous = self.api.token
        self.api.token = token
        try:
            self.api.cikis()
        except ApiError:
            self.api.token = ""
        self.api.token = previous

    def _loop(self, generation):
        if generation != self._generation or not self.winfo_exists():
            return
        self._refresh(generation)
        self.after(4000, lambda: self._loop(generation))

    def _refresh(self, generation=None):
        generation = self._generation if generation is None else generation
        page = self.page

        def work():
            try:
                self.api.beat(device_snapshot())
                data = self._load(page)
                error = ""
            except ApiError as exc:
                data, error = None, str(exc)
            try:
                self.after(0, lambda: self._apply(generation, page, data, error))
            except tk.TclError:
                pass

        threading.Thread(target=work, daemon=True).start()

    def _load(self, page):
        if page == "dagilim":
            return {"dagilim": self.api.distribution(), "talepler": self.api.requests()}
        if page == "talepler":
            return {"talepler": self.api.requests()}
        if page == "dosyalar":
            return {"dosyalar": self.api.files()}
        if page == "mesajlar":
            return {"mesajlar": self.api.messages()}
        return {"bilgisayarlar": self.api.computers()}

    def _apply(self, generation, page, data, error):
        try:
            alive = self.winfo_exists()
        except tk.TclError:
            return
        if not alive or generation != self._generation:
            return
        if hasattr(self, "link") and self.link.winfo_exists():
            self.link.configure(text=error or "Sunucu bağlı", fg="#F0B4B7" if error else "#8FDBA8")
        if error == "Oturum kapandı. Yeniden gir.":
            self._logout()
            return
        if self.page != page or self._busy or not hasattr(self, "board") or not self.board.winfo_exists():
            return
        for child in self.board.winfo_children():
            child.destroy()
        if error:
            tk.Label(self.board, text=error, bg=CANVAS, fg=DANGER, font=(FONT, 12)).pack(anchor="w")
            return
        if page == "dagilim":
            self._paint_spread(data)
        elif page == "talepler":
            self._paint_requests(data["talepler"])
        elif page == "dosyalar":
            self._paint_files(data["dosyalar"])
        elif page == "mesajlar":
            self._paint_messages(data["mesajlar"])
        else:
            self._paint_computers(data["bilgisayarlar"])

    def _show(self, name):
        self.page = name
        self._paint_nav()
        for child in self.content.winfo_children():
            child.destroy()
        titles = {
            "dagilim": ("Talep dağılımı", "Kimin üzerinde kaç açık talep var. Yeni talebi kişiye buradan verirsin."),
            "talepler": ("Talepler", "Sana gelen ve senin açtığın talepler. Yönetici hepsini görür."),
            "dosyalar": ("Dosyalar", "Paylaşılan dosya sunucuda durur. Bu listedeki herkes indirir."),
            "mesajlar": ("Mesajlar", "Ağa bağlı bilgisayarlar aynı mesaj tahtasını görür."),
            "bilgisayarlar": ("Bilgisayarlar", "Her kurulum kendi bilgisayarını ayrı gösterir."),
        }
        title, note = titles[name]
        head = tk.Frame(self.content, bg=CANVAS)
        head.pack(fill="x", padx=20, pady=(16, 2))
        tk.Label(head, text=title, bg=CANVAS, fg=INK, font=(FONT, 22, "bold")).pack(side="left")
        if name == "dagilim":
            Click(head, "Talep ver", self._new_request, ACCENT, "white", ACCENT_HOVER).pack(side="right")
        elif name == "dosyalar":
            Click(head, "Dosya yükle", self._upload, ACCENT, "white", ACCENT_HOVER).pack(side="right")
        elif name == "bilgisayarlar" and self.user and self.user.get("rol") == "Yönetici":
            Click(head, "Kullanıcı ekle", self._new_user, ACCENT, "white", ACCENT_HOVER).pack(side="right")
        tk.Label(self.content, text=note, bg=CANVAS, fg=MUTED, font=(FONT, 12), wraplength=760, justify="left").pack(anchor="w", padx=20)
        if name == "mesajlar":
            composer = tk.Frame(self.content, bg=CANVAS)
            composer.pack(side="bottom", fill="x", padx=20, pady=(0, 16))
            text = tk.StringVar()
            tk.Entry(composer, textvariable=text, relief="flat", highlightthickness=1, highlightbackground=LINE, font=(FONT, 13)).pack(side="left", fill="x", expand=True, ipady=6)
            Click(composer, "Gönder", lambda: self._send(text), ACCENT, "white", ACCENT_HOVER).pack(side="left", padx=(8, 0))
        self.board = tk.Frame(self.content, bg=CANVAS)
        self.board.pack(fill="both", expand=True, padx=20, pady=(10, 16))
        tk.Label(self.board, text="Yükleniyor", bg=CANVAS, fg=MUTED, font=(FONT, 12)).pack(anchor="w")
        self._refresh()

    def _paint_spread(self, data):
        row = tk.Frame(self.board, bg=CANVAS)
        row.pack(fill="x")
        for person in data["dagilim"]:
            card = tk.Frame(row, bg=CARD, highlightbackground=LINE, highlightthickness=1, padx=12, pady=10)
            card.pack(side="left", fill="x", expand=True, padx=(0, 8))
            tk.Label(card, text=person["ad"], bg=CARD, fg=INK, font=(FONT, 13, "bold")).pack(anchor="w")
            tk.Label(card, text=person["rol"], bg=CARD, fg=MUTED, font=(FONT, 11)).pack(anchor="w")
            count = f"{person['acik']} açık" if person["acik"] else "Boşta"
            tk.Label(card, text=count, bg=CARD, fg=ACCENT if person["acik"] else OK, font=(FONT, 16, "bold")).pack(anchor="w", pady=(6, 8))
            Click(card, "Talep ver", lambda item=person: self._new_request(item), SOFT, ACCENT, HOVER, padx=8, pady=4).pack(anchor="w")
        mine = [item for item in data["talepler"] if item["alan_id"] == self.user.get("id") and item["durum"] != "Tamamlandı"]
        tk.Label(self.board, text="Sana gelen açık talepler", bg=CANVAS, fg=INK, font=(FONT, 14, "bold")).pack(anchor="w", pady=(16, 6))
        if not mine:
            tk.Label(self.board, text="Açık talebin yok.", bg=CANVAS, fg=MUTED, font=(FONT, 12)).pack(anchor="w")
            return
        self._paint_requests(mine)

    def _paint_requests(self, rows):
        if not rows:
            tk.Label(self.board, text="Talep yok.", bg=CANVAS, fg=MUTED, font=(FONT, 12)).pack(anchor="w")
            return
        for item in rows:
            card = tk.Frame(self.board, bg=CARD, highlightbackground=LINE, highlightthickness=1, padx=12, pady=8)
            card.pack(fill="x", pady=3)
            tk.Label(card, text=f"{item['no']}  ·  {item['baslik']}", bg=CARD, fg=INK, font=(FONT, 13, "bold")).pack(anchor="w")
            tk.Label(card, text=f"{item['acan']} → {item['alan']}  ·  {item['zaman']}", bg=CARD, fg=MUTED, font=(FONT, 11)).pack(anchor="w")
            if item["not"]:
                tk.Label(card, text=item["not"], bg=CARD, fg=INK, font=(FONT, 12), wraplength=680, justify="left").pack(anchor="w", pady=(2, 0))
            actions = tk.Frame(card, bg=CARD)
            actions.pack(anchor="w", pady=(6, 0))
            tone = {"Açık": ACCENT, "Üzerinde": "#9A6B12", "Tamamlandı": OK}.get(item["durum"], INK)
            tk.Label(actions, text=item["durum"], bg=CARD, fg=tone, font=(FONT, 12, "bold")).pack(side="left", padx=(0, 8))
            for status in STATUSES:
                if status == item["durum"]:
                    continue
                Click(actions, status, lambda row=item, value=status: self._move(row, value), SOFT, ACCENT, HOVER, padx=8, pady=3).pack(side="left", padx=(0, 4))

    def _paint_files(self, rows):
        if not rows:
            tk.Label(self.board, text="Henüz dosya yok.", bg=CANVAS, fg=MUTED, font=(FONT, 12)).pack(anchor="w")
            return
        for item in rows:
            card = tk.Frame(self.board, bg=CARD, highlightbackground=LINE, highlightthickness=1, padx=12, pady=8)
            card.pack(fill="x", pady=3)
            tk.Label(card, text=item["ad"], bg=CARD, fg=INK, font=(FONT, 13, "bold")).pack(side="left")
            tk.Label(card, text=f"{item['sahip']}  ·  {file_size(item['boy'])}  ·  {item['zaman']}", bg=CARD, fg=MUTED, font=(FONT, 11)).pack(side="left", padx=12)
            Click(card, "İndir", lambda row=item: self._download(row), SOFT, ACCENT, HOVER, padx=8, pady=3).pack(side="right")

    def _paint_messages(self, rows):
        if not rows:
            tk.Label(self.board, text="Tahta boş.", bg=CANVAS, fg=MUTED, font=(FONT, 12)).pack(anchor="w")
            return
        for item in rows:
            card = tk.Frame(self.board, bg=CARD, highlightbackground=LINE, highlightthickness=1, padx=12, pady=8)
            card.pack(fill="x", pady=3)
            tk.Label(card, text=f"{item['ad']}  ·  {item['zaman']}", bg=CARD, fg=MUTED, font=(FONT, 11)).pack(anchor="w")
            tk.Label(card, text=item["metin"], bg=CARD, fg=INK, font=(FONT, 13), wraplength=700, justify="left").pack(anchor="w")

    def _paint_computers(self, rows):
        if not rows:
            tk.Label(self.board, text="Henüz bağlanan bilgisayar yok. İstemci açık kalınca burada görünür.", bg=CANVAS, fg=MUTED, font=(FONT, 12), wraplength=680, justify="left").pack(anchor="w")
            return
        for item in rows:
            card = tk.Frame(self.board, bg=CARD, highlightbackground=LINE, highlightthickness=1, padx=12, pady=8)
            card.pack(fill="x", pady=3)
            mark = "Çevrimiçi" if item["cevrimici"] else "Çevrimdışı"
            tk.Label(card, text=item["cihaz"] or "Bilgisayar", bg=CARD, fg=INK, font=(FONT, 13, "bold")).pack(side="left")
            tk.Label(card, text=f"{item['ad']}  ·  {item['ip'] or 'IP yok'}  ·  {item['sistem']}", bg=CARD, fg=MUTED, font=(FONT, 11)).pack(side="left", padx=12)
            tk.Label(card, text=mark, bg=CARD, fg=OK if item["cevrimici"] else MUTED, font=(FONT, 12, "bold")).pack(side="right")

    def _move(self, row, status):
        def work():
            try:
                self.api.set_status(row["id"], status)
            except ApiError as exc:
                self.after(0, lambda: messagebox.showwarning(APP, str(exc), parent=self))
                return
            self.after(0, self._refresh)

        threading.Thread(target=work, daemon=True).start()

    def _new_request(self, person=None):
        self._busy = True

        def work():
            try:
                users = self.api.users()
            except ApiError as exc:
                self._busy = False
                self.after(0, lambda: messagebox.showwarning(APP, str(exc), parent=self))
                return
            self.after(0, lambda: self._request_dialog(users, person))

        threading.Thread(target=work, daemon=True).start()

    def _request_dialog(self, users, person):
        dialog = tk.Toplevel(self)
        dialog.title("Talep ver")
        dialog.configure(bg=CANVAS)
        dialog.transient(self)
        dialog.resizable(False, False)
        card = tk.Frame(dialog, bg=CARD, padx=16, pady=14, highlightbackground=LINE, highlightthickness=1)
        card.pack(padx=12, pady=12)
        tk.Label(card, text="Talep ver", bg=CARD, fg=INK, font=(FONT, 16, "bold")).pack(anchor="w")
        title = tk.StringVar()
        note = tk.StringVar()
        labels = [f"{item['ad']}  ·  {item['kullanici']}" for item in users]
        picked = labels[0] if labels else ""
        if person is not None:
            for label, item in zip(labels, users):
                if item["id"] == person["id"]:
                    picked = label
        target = tk.StringVar(value=picked)
        for caption, var in (("Başlık", title), ("Not", note)):
            tk.Label(card, text=caption, bg=CARD, fg=MUTED, font=(FONT, 11, "bold")).pack(anchor="w", pady=(8, 2))
            tk.Entry(card, textvariable=var, relief="flat", highlightthickness=1, highlightbackground=LINE, font=(FONT, 13), width=42).pack(fill="x", ipady=5)
        tk.Label(card, text="Kime", bg=CARD, fg=MUTED, font=(FONT, 11, "bold")).pack(anchor="w", pady=(8, 2))
        ttk.Combobox(card, textvariable=target, values=labels, state="readonly", font=(FONT, 12), width=40).pack(fill="x", ipady=4)

        def close():
            self._busy = False
            dialog.destroy()

        def save():
            if target.get() not in labels:
                messagebox.showwarning(APP, "Kişi seç.", parent=dialog)
                return
            assignee = users[labels.index(target.get())]

            def work():
                try:
                    self.api.create_request(title.get(), note.get(), assignee["id"])
                except ApiError as exc:
                    self.after(0, lambda: messagebox.showwarning(APP, str(exc), parent=dialog))
                    return
                self.after(0, close)
                self.after(0, self._refresh)

            threading.Thread(target=work, daemon=True).start()

        Click(card, "Gönder", save, ACCENT, "white", ACCENT_HOVER).pack(anchor="e", pady=(14, 0))
        dialog.protocol("WM_DELETE_WINDOW", close)
        dialog.bind("<Escape>", lambda _e: close())
        dialog.grab_set()

    def _new_user(self):
        self._busy = True
        dialog = tk.Toplevel(self)
        dialog.title("Kullanıcı ekle")
        dialog.configure(bg=CANVAS)
        dialog.transient(self)
        dialog.resizable(False, False)
        card = tk.Frame(dialog, bg=CARD, padx=16, pady=14, highlightbackground=LINE, highlightthickness=1)
        card.pack(padx=12, pady=12)
        tk.Label(card, text="Kullanıcı ekle", bg=CARD, fg=INK, font=(FONT, 16, "bold")).pack(anchor="w")
        name = tk.StringVar()
        username = tk.StringVar()
        password = tk.StringVar()
        role = tk.StringVar(value=ROLES[-1])
        for caption, var, secret in (("Ad soyad", name, False), ("Kullanıcı adı", username, False), ("Parola", password, True)):
            tk.Label(card, text=caption, bg=CARD, fg=MUTED, font=(FONT, 11, "bold")).pack(anchor="w", pady=(8, 2))
            tk.Entry(card, textvariable=var, show="•" if secret else "", relief="flat", highlightthickness=1, highlightbackground=LINE, font=(FONT, 13), width=42).pack(fill="x", ipady=5)
        tk.Label(card, text="Rol", bg=CARD, fg=MUTED, font=(FONT, 11, "bold")).pack(anchor="w", pady=(8, 2))
        ttk.Combobox(card, textvariable=role, values=list(ROLES), state="readonly", font=(FONT, 12), width=40).pack(fill="x", ipady=4)

        def close():
            self._busy = False
            dialog.destroy()

        def save():
            def work():
                try:
                    self.api.add_user(name.get(), username.get(), password.get(), role.get())
                except ApiError as exc:
                    self.after(0, lambda: messagebox.showwarning(APP, str(exc), parent=dialog))
                    return
                self.after(0, close)

            threading.Thread(target=work, daemon=True).start()

        Click(card, "Kaydet", save, ACCENT, "white", ACCENT_HOVER).pack(anchor="e", pady=(14, 0))
        dialog.protocol("WM_DELETE_WINDOW", close)
        dialog.bind("<Escape>", lambda _e: close())
        dialog.grab_set()

    def _upload(self):
        picked = filedialog.askopenfilename(parent=self, filetypes=[("Paylaşılacak dosya", " ".join(f"*{ext}" for ext in sorted(FILE_EXT)))])
        if not picked:
            return

        def work():
            try:
                self.api.upload(picked)
            except (ApiError, OSError, ValueError) as exc:
                self.after(0, lambda: messagebox.showwarning(APP, str(exc), parent=self))
                return
            self.after(0, self._refresh)

        threading.Thread(target=work, daemon=True).start()

    def _download(self, row):
        target = filedialog.asksaveasfilename(parent=self, initialfile=row["ad"])
        if not target:
            return

        def work():
            try:
                payload = self.api.download(row["id"])
                Path(target).write_bytes(payload)
            except (ApiError, OSError) as exc:
                self.after(0, lambda: messagebox.showwarning(APP, str(exc), parent=self))
                return
            self.after(0, lambda: messagebox.showinfo(APP, "Dosya indirildi.", parent=self))

        threading.Thread(target=work, daemon=True).start()

    def _send(self, text):
        message = text.get()

        def work():
            try:
                self.api.send_message(message)
            except ApiError as exc:
                self.after(0, lambda: messagebox.showwarning(APP, str(exc), parent=self))
                return
            self.after(0, lambda: text.set(""))
            self.after(0, self._refresh)

        threading.Thread(target=work, daemon=True).start()


def _self_check():
    import tempfile
    import time
    backend = Path(__file__).resolve().parents[1] / "vkv-backend"
    sys.path.insert(0, str(backend))
    from vkv_backend.http import Server
    from vkv_backend.hub import Hub
    folder = tempfile.TemporaryDirectory()
    client_root = Path(folder.name) / "istemci"
    client_root.mkdir()
    global istemci_dir
    previous_dir = istemci_dir
    istemci_dir = lambda: client_root
    hub = Hub(Path(folder.name))
    httpd = Server(("127.0.0.1", 0), hub)
    thread = threading.Thread(target=httpd.serve_forever, daemon=True)
    thread.start()
    port = httpd.server_address[1]
    try:
        api = Api()
        api.client_id = "istemci-ekran"
        api.set_base(f"127.0.0.1:{port}")
        assert api.durum()["kurulum"]
        api.kurulum(hub.setup_key, "Furkan Hamarat", "furkan", "parola1")
        user = api.giris("furkan", "parola1")
        other = api.add_user("Ali Veli", "ali", "parola2", "Kullanıcı")
        api.create_request("Fatura kontrol", "Mart", other["id"])
        app = AgApp(start=False)
        app.withdraw()
        app.api = api
        app._enter(user)
        app.update()
        for page in ("dagilim", "talepler", "dosyalar", "mesajlar", "bilgisayarlar"):
            app._show(page)
            app.update()
        time.sleep(0.4)
        app.update()
        texts = []

        def walk(widget):
            if isinstance(widget, tk.Label):
                texts.append(widget.cget("text"))
            for child in widget.winfo_children():
                walk(child)

        walk(app)
        joined = "\n".join(texts)
        assert "Talep dağılımı" in joined or "Bilgisayarlar" in joined
        app._generation += 1
        app.destroy()
    finally:
        istemci_dir = previous_dir
        httpd.shutdown()
        httpd.server_close()
        hub.close()
        folder.cleanup()
    print("ekran-ok")


def main():
    if "--check" in sys.argv:
        _self_check()
        return
    AgApp().mainloop()


if __name__ == "__main__":
    main()
