#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
İş Merkezi — iş takibi, fatura önizleme, ortak klasör ve isteğe bağlı e-posta.

Veriler bu bilgisayarda kalır. USB paketinde veriler uygulamanın yanındaki
Veriler klasöründedir. E-posta parolası veritabanına yazılmaz; Mac'te
Anahtar Zinciri'nde, Windows'ta Kimlik Yöneticisi'nde durur. Ekler çalıştırılmaz.
"""
import base64
import calendar
import ctypes
import email
import email.errors
import email.policy
import hashlib
import html
import imaplib
import json
import math
import os
import posixpath
import queue
import re
import secrets
import shutil
import sqlite3
import subprocess
import sys
import tempfile
import threading
import urllib.error
import urllib.parse
import urllib.request
import webbrowser
import zipfile
from datetime import date, datetime, timedelta, timezone
from email.header import decode_header, make_header
from pathlib import Path
from xml.etree import ElementTree as ET
import tkinter as tk
import tkinter.font as tkfont
from tkinter import filedialog, messagebox, simpledialog, ttk

APP_NAME = "VKV İş Takip Merkezi"
if sys.platform == "win32":
    FONT = "Segoe UI"
elif sys.platform == "darwin":
    FONT = ".AppleSystemUIFont"
else:
    FONT = "Helvetica Neue"
DATE_FMT = "%Y-%m-%d"
KEYCHAIN_SERVICE = "com.furkan.ismerkezi.mail"
MAX_BYTES = 20 * 1024 * 1024
MAX_SCAN_FILES = 400
MAIL_WINDOW_DAYS = 21
MAIL_MESSAGE_CAP = 30
LOCAL_MAIL_HOST = "local.mail"

def _data_dir():
    portable = os.environ.get("IS_MERKEZI_HOME", "").strip().strip('"')
    if portable:
        return Path(portable)
    if sys.platform == "darwin":
        return Path.home() / "Library" / "Application Support" / "IsMerkezi"
    if sys.platform == "win32":
        return Path(os.environ.get("APPDATA", str(Path.home()))) / "IsMerkezi"
    return Path.home() / ".is_merkezi"


DATA_DIR = _data_dir()
ATTACH_DIR = DATA_DIR / "Ekler"
LEDGER_DIR = ATTACH_DIR / "Kasa"
REPORT_DIR = ATTACH_DIR / "Raporlar"
OFFER_DIR = ATTACH_DIR / "Teklifler"
PREVIEW_DIR = DATA_DIR / "Onizleme"
NOTE_DIR = DATA_DIR / "Notlar"
DB_PATH = DATA_DIR / "is_merkezi.db"

NAV = "#1B2437"
NAV_ON = "#2C3B52"
NAV_HOVER = "#243248"
NAV_TEXT = "#C5D0DE"
NAV_MUTED = "#8E9BB0"
ACCENT = "#3B6CF5"
ACCENT_HOVER = "#2F5AE0"
CANVAS = "#F4F6FB"
INK = "#1A2332"
MUTED = "#8A94A6"
CARD = "#FFFFFF"
LINE = "#E6EAF2"
SOFT = "#EEF3FF"
SOFT_HOVER = "#E3EBFF"
HOVER = "#F3F6FB"
PAPER = "#F7F8FC"
DANGER = "#E5484D"
DANGER_SOFT = "#FDECEC"
WARN = "#D9920B"
WARN_SOFT = "#FFF6DE"
OK = "#1F9D55"
OK_SOFT = "#E7F6EE"

PAGES = ["İş takibi", "Görevler", "Faturalar", "Hatırlatmalar", "Firmalar", "Gelir-Gider", "Projeler"]
PAGE_TITLES = {"Firmalar": "Rehber"}
ALLOWED_EXT = {".pdf", ".png", ".jpg", ".jpeg", ".webp", ".tif", ".tiff", ".heic"}
NUMERIC_COLS = {"E", "H", "I", "K", "M"}
EXCEL_HEADERS = [
    "Fatura No (Klasör)",
    "Klasör / Dosya Yolu",
    "Dosya Adı",
    "Belge Fatura No",
    "Fatura Tarihi",
    "Satıcı Unvanı",
    "Satıcı VKN/TCKN",
    "Matrah (KDV Hariç)",
    "KDV Tutarı",
    "Para Birimi",
    "Genel Toplam",
    "Durum",
    "Aktarım Tarihi",
]
MONTHS = "Ocak Şubat Mart Nisan Mayıs Haziran Temmuz Ağustos Eylül Ekim Kasım Aralık".split()
PRIORITY_RANK = {"Acil": 0, "Yüksek": 1, "Normal": 2, "Düşük": 3}
PRIORITY_CELL = {
    "Acil": (DANGER_SOFT, DANGER),
    "Yüksek": ("#FFF1E4", "#E07A1F"),
    "Normal": (SOFT, ACCENT),
    "Düşük": ("#F3EEFF", "#7C5CBF"),
}
TASK_STRIPE = ((CARD, INK), (PAPER, INK))
STATUS_PILL = {
    "Bekliyor": (WARN_SOFT, WARN),
    "Devam ediyor": (SOFT, ACCENT),
    "Tamamlandı": (OK_SOFT, OK),
    "İptal": (PAPER, MUTED),
    "Ödenmedi": (DANGER_SOFT, DANGER),
    "Kısmi ödendi": (WARN_SOFT, WARN),
    "Ödendi": (OK_SOFT, OK),
    "Aktif": (SOFT, ACCENT),
}

SCHEMAS = {
    "Görevler": {
        "table": "tasks",
        "title": "Görev",
        "fields": [
            ("Başlık", "title", "text", True),
            ("Tür", "issue_type", "choice", ["Görev", "Hata", "İyileştirme"]),
            ("Açıklama", "description", "multiline", False),
            ("Sorumlu", "assignee", "text", False),
            ("Öncelik", "priority", "choice", ["Düşük", "Normal", "Yüksek", "Acil"]),
            ("Durum", "status", "choice", ["Bekliyor", "Devam ediyor", "Tamamlandı", "İptal"]),
            ("Başlangıç tarihi", "start_date", "date", False),
            ("Son tarih", "due_date", "date", False),
            ("Tekrar", "repeat", "choice", ["Yok", "Her gün", "Her hafta", "Her ay"]),
            ("Etiket", "tag", "text", False),
        ],
        "columns": [
            ("issue_key", "Anahtar"),
            ("title", "Görev"),
            ("issue_type", "Tür"),
            ("priority", "Öncelik"),
            ("status", "Durum"),
            ("assignee", "Sorumlu"),
            ("due_date", "Son tarih"),
            ("project_name", "Proje"),
            ("tag", "Etiket"),
        ],
    },
    "Faturalar": {
        "table": "invoices",
        "title": "Fatura",
        "fields": [
            ("Durum", "status", "choice", ["Ödenmedi", "Ödendi", "Kısmi ödendi", "İptal"]),
            ("Tür", "kind", "choice", ["Gelen", "Kesilen"]),
            ("Firma / kişi", "party", "text", True),
            ("Fatura no", "invoice_no", "text", False),
            ("Açıklama", "description", "text", False),
            ("Tutar (₺)", "amount", "number", False),
            ("Fatura tarihi", "invoice_date", "date", False),
            ("Vade tarihi", "due_date", "date", False),
            ("Dosya", "attachment", "file", False),
        ],
        "columns": [
            ("kind", "Tür"),
            ("party", "Firma / kişi"),
            ("invoice_no", "Fatura no"),
            ("amount", "Tutar (₺)"),
            ("due_date", "Vade"),
            ("status", "Durum"),
        ],
    },
    "Hatırlatmalar": {
        "table": "reminders",
        "title": "Hatırlatma",
        "fields": [
            ("Başlık", "title", "text", True),
            ("Açıklama", "description", "multiline", False),
            ("Tarih", "remind_date", "date", True),
            ("Saat (HH:MM)", "remind_time", "time", False),
            ("Tekrar", "repeat", "choice", ["Yok", "Her gün", "Her hafta", "Her ay"]),
            ("Durum", "status", "choice", ["Aktif", "Tamamlandı"]),
        ],
        "columns": [
            ("title", "Hatırlatma"),
            ("remind_date", "Tarih"),
            ("remind_time", "Saat"),
            ("repeat", "Tekrar"),
            ("status", "Durum"),
        ],
    },
    "Firmalar": {
        "table": "contacts",
        "title": "Firma / kişi",
        "fields": [
            ("Firma / kişi adı", "name", "text", True),
            ("Yetkili kişi", "contact_person", "text", False),
            ("Telefon", "phone", "text", False),
            ("E-posta", "email", "text", False),
            ("Vergi / kimlik no", "tax_no", "text", False),
            ("Adres", "address", "multiline", False),
            ("Notlar", "notes", "multiline", False),
        ],
        "columns": [
            ("record_kind", "Kayıt"),
            ("scope", "Kapsam"),
            ("name", "Ad"),
            ("contact_person", "Firma / yetkili"),
            ("extension", "Dahili"),
            ("phone", "Telefon"),
            ("email", "E-posta"),
        ],
    },
    "Gelir-Gider": {
        "table": "transactions",
        "title": "Gelir / gider kaydı",
        "fields": [
            ("Tür", "kind", "choice", ["Gelir", "Gider"]),
            ("Kategori", "category", "category", True),
            ("Açıklama", "description", "text", False),
            ("Tutar (₺)", "amount", "number", True),
            ("Tarih", "trans_date", "date", True),
            ("İlgili firma / kişi", "party", "text", False),
            ("Ödeme yöntemi", "payment_method", "choice", ["Nakit", "Banka", "Kredi kartı", "Diğer"]),
        ],
        "columns": [
            ("kind", "Tür"),
            ("category", "Kategori"),
            ("description", "Açıklama"),
            ("amount", "Tutar (₺)"),
            ("trans_date", "Tarih"),
            ("party", "Firma / kişi"),
        ],
    },
    "Projeler": {
        "table": "projects",
        "title": "Proje",
        "fields": [
            ("Proje adı", "name", "text", True),
            ("Anahtar", "proj_key", "text", False),
            ("Açıklama", "description", "multiline", False),
            ("Başlangıç tarihi", "start_date", "date", False),
            ("Bitiş tarihi", "end_date", "date", False),
            ("Durum", "status", "choice", ["Planlandı", "Devam ediyor", "Beklemede", "Tamamlandı", "İptal"]),
            ("Bütçe (₺)", "budget", "number", False),
            ("Notlar", "notes", "multiline", False),
        ],
        "columns": [
            ("name", "Proje"),
            ("proj_key", "Anahtar"),
            ("start_date", "Başlangıç"),
            ("end_date", "Bitiş"),
            ("status", "Durum"),
            ("budget", "Bütçe (₺)"),
        ],
    },
}

TABLES = {schema["table"] for schema in SCHEMAS.values()} | {"notify_log", "settings", "ignored_sources", "reports", "offers"}
DETAIL_KEYS = ("description", "notes", "address", "attachment", "excel_status", "source")
MONEY_KEYS = {"amount", "budget"}
SAFE_KEY = re.compile(r"^[a-z_]{1,40}$")
SETTING_KEY = re.compile(r"^[a-z0-9_]{1,40}$")


def today():
    return date.today().strftime(DATE_FMT)


def long_today():
    day = date.today()
    return f"{day.day} {MONTHS[day.month - 1]} {day.year}"


def parse_day(value):
    if not value:
        return None
    try:
        return datetime.strptime(str(value).strip(), DATE_FMT).date()
    except ValueError:
        return None


def excel_serial(day):
    return (day - date(1899, 12, 30)).days


def from_excel_serial(value):
    try:
        return (date(1899, 12, 30) + timedelta(days=int(float(value)))).strftime(DATE_FMT)
    except (TypeError, ValueError, OverflowError):
        return ""


def parse_amount(value, required):
    text = str(value or "").strip().replace(" ", "").replace("₺", "").replace("TL", "").replace("tl", "")
    if not text:
        if required:
            raise ValueError("Tutar gerekli.")
        return None
    if re.fullmatch(r"\d{1,3}(\.\d{3})+", text):
        text = text.replace(".", "")
    elif re.fullmatch(r"\d{1,3}(\.\d{3})+,\d{1,2}", text):
        text = text.replace(".", "").replace(",", ".")
    elif "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        text = text.replace(",", ".")
    try:
        return round(float(text), 2)
    except ValueError as exc:
        raise ValueError("Tutar sayı olmalı.") from exc


def format_amount(value):
    if value is None or value == "":
        return ""
    try:
        number = float(value)
    except (TypeError, ValueError):
        return str(value)
    text = f"{number:,.2f}"
    return text.replace(",", "X").replace(".", ",").replace("X", ".")


def party_initials(name):
    parts = [part for part in str(name or "").split() if part]
    if not parts:
        return "—"
    if len(parts) == 1:
        return parts[0][:2].upper()
    return (parts[0][0] + parts[-1][0]).upper()


def invoice_money(row):
    try:
        total = float(row["amount"] or 0)
    except (TypeError, ValueError):
        total = 0.0
    try:
        paid = float(row["paid_amount"] or 0) if "paid_amount" in row.keys() else 0.0
    except (TypeError, ValueError):
        paid = 0.0
    status = (row["status"] or "") if "status" in row.keys() else ""
    if status == "Ödendi":
        paid = max(paid, total)
    elif status in ("Ödenmedi", "İptal"):
        paid = 0.0
    if total:
        paid = min(max(paid, 0.0), total)
    return total, max(paid, 0.0)


def invoice_events(row):
    raw = row["events_json"] if "events_json" in row.keys() else ""
    try:
        data = json.loads(raw or "[]")
    except json.JSONDecodeError:
        return []
    return [item for item in data if isinstance(item, dict)]


def phone_digits(value):
    digits = "".join(ch for ch in str(value or "") if ch.isdigit())
    if digits.startswith("00"):
        digits = digits[2:]
    if digits.startswith("0"):
        digits = "90" + digits[1:]
    elif len(digits) == 10:
        digits = "90" + digits
    return digits


OUTLOOK_CLOSED = "Outlook açık değil. Arama yalnız açık Outlook adres defterinde yapılır."
OUTLOOK_MAC_SCRIPT = r'''
on clean(value)
    set raw to ""
    try
        set raw to value as text
    end try
    set AppleScript's text item delimiters to {tab, return, linefeed}
    set bits to text items of raw
    set AppleScript's text item delimiters to " "
    set raw to bits as text
    set AppleScript's text item delimiters to ""
    if length of raw > 120 then set raw to text 1 thru 120 of raw
    return raw
end clean

on addHit(displayName, companyName, emailAddress, phoneNumber, collected)
    if displayName is "" then return collected
    set lineText to "Outlook" & tab & clean(displayName) & tab & clean(companyName) & tab & clean(emailAddress) & tab & clean(phoneNumber)
    if collected is "" then return lineText
    return collected & linefeed & lineText
end addHit

on run argv
    set q to item 1 of argv
    set collected to ""
    set foundCount to 0
    tell application "Microsoft Outlook"
        if not running then return "CLOSED"
        set hits to contacts whose display name contains q or company contains q
        repeat with c in hits
            if foundCount > 19 then exit repeat
            set displayName to ""
            set companyName to ""
            set emailAddress to ""
            set phoneNumber to ""
            try
                set displayName to display name of c
            end try
            try
                set companyName to company of c
            end try
            try
                set emailAddress to address of first email address of c
            end try
            try
                set phoneNumber to phone number of first phone number of c
            end try
            set collected to my addHit(displayName, companyName, emailAddress, phoneNumber, collected)
            set foundCount to foundCount + 1
        end repeat
    end tell
    return collected
end run
'''
OUTLOOK_WINDOWS_SCRIPT = r'''
$q = $env:OUTLOOK_QUERY
if ([string]::IsNullOrWhiteSpace($q)) { exit 0 }
function Clean([string]$value) {
    if (-not $value) { return "" }
    $value = ($value -replace "[\t\r\n]", " ").Trim()
    if ($value.Length -gt 120) { $value = $value.Substring(0, 120) }
    return $value
}
$seen = New-Object System.Collections.Generic.HashSet[string]
function Emit($name, $company, $email, $phone) {
    $name = Clean $name
    if (-not $name) { return }
    $key = ($name + "|" + (Clean $email)).ToLower()
    if (-not $seen.Add($key)) { return }
    if ($seen.Count -gt 20) { return }
    Write-Output ("Outlook`t" + $name + "`t" + (Clean $company) + "`t" + (Clean $email) + "`t" + (Clean $phone))
}
try {
    $outlook = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
} catch {
    [Console]::Error.WriteLine("CLOSED")
    exit 2
}
$ns = $outlook.GetNamespace("MAPI")
$safe = $q.Replace("'", "''").Replace("%", "").Replace("_", "")
$filter = "@SQL=""urn:schemas:contacts:cn"" LIKE '%$safe%'"
try {
    $items = $ns.GetDefaultFolder(10).Items.Restrict($filter)
    foreach ($item in $items) {
        if ($seen.Count -ge 20) { break }
        $phone = $item.BusinessTelephoneNumber
        if (-not $phone) { $phone = $item.MobileTelephoneNumber }
        Emit $item.FullName $item.CompanyName $item.Email1Address $phone
    }
} catch {}
try {
    $recipient = $ns.CreateRecipient($q)
    if ($recipient.Resolve()) {
        $entry = $recipient.AddressEntry
        $name = $entry.Name
        $email = ""
        $company = ""
        try {
            $user = $entry.GetExchangeUser()
            if ($null -ne $user) {
                if ($user.Name) { $name = $user.Name }
                $email = $user.PrimarySmtpAddress
                $company = $user.CompanyName
            }
        } catch {}
        if (-not $email) {
            try { $email = $entry.Address } catch {}
        }
        Emit $name $company $email ""
    }
} catch {}
'''


def _outlook_open():
    if sys.platform == "darwin":
        if not Path("/Applications/Microsoft Outlook.app").exists():
            return False
        try:
            done = subprocess.run(
                ["osascript", "-e", 'application "Microsoft Outlook" is running'],
                capture_output=True,
                text=True,
                timeout=5,
            )
        except (OSError, subprocess.TimeoutExpired):
            return False
        return done.returncode == 0 and done.stdout.strip().lower() == "true"
    if sys.platform == "win32":
        try:
            done = subprocess.run(
                ["tasklist", "/FI", "IMAGENAME eq OUTLOOK.EXE", "/NH"],
                capture_output=True,
                text=True,
                timeout=8,
            )
        except (OSError, subprocess.TimeoutExpired):
            return False
        return "OUTLOOK.EXE" in (done.stdout or "").upper()
    return False


def _parse_outlook_hits(stdout):
    hits = []
    for line in (stdout or "").splitlines():
        parts = line.split("\t")
        if len(parts) < 5 or not parts[1].strip():
            continue
        hits.append({
            "source": "Outlook",
            "name": parts[1].strip(),
            "company": parts[2].strip(),
            "email": parts[3].strip(),
            "phone": parts[4].strip(),
        })
    return hits[:20]


def lookup_address_book(query):
    text = " ".join(str(query or "").split())[:60]
    if len(text) < 2:
        return [], ""
    if sys.platform not in {"darwin", "win32"}:
        return [], "Outlook araması Mac ve Windows içindir."
    if not _outlook_open():
        return [], OUTLOOK_CLOSED
    try:
        if sys.platform == "darwin":
            done = subprocess.run(
                ["osascript", "-e", OUTLOOK_MAC_SCRIPT, text],
                capture_output=True,
                text=True,
                timeout=20,
            )
        else:
            env = os.environ.copy()
            env["OUTLOOK_QUERY"] = text
            done = subprocess.run(
                ["powershell.exe", "-NoProfile", "-STA", "-NonInteractive", "-Command", OUTLOOK_WINDOWS_SCRIPT],
                capture_output=True,
                text=True,
                timeout=20,
                env=env,
            )
    except subprocess.TimeoutExpired:
        return [], "Outlook yanıt vermedi."
    err = (done.stderr or "").lower()
    if "closed" in err:
        return [], OUTLOOK_CLOSED
    if done.returncode != 0:
        if "not authorized" in err or "-1743" in err:
            return [], "Outlook izni yok. Sistem Ayarları, Gizlilik ve Güvenlik, Otomasyon bölümünden izin ver."
        return [], "Outlook adres defteri açılamadı."
    if (done.stdout or "").strip() == "CLOSED":
        return [], OUTLOOK_CLOSED
    return _parse_outlook_hits(done.stdout), ""


def fold(value):
    text = str(value or "").replace("\u0307", "")
    text = text.translate(str.maketrans({
        "İ": "i", "I": "i", "ı": "i",
        "Ş": "s", "ş": "s",
        "Ğ": "g", "ğ": "g",
        "Ü": "u", "ü": "u",
        "Ö": "o", "ö": "o",
        "Ç": "c", "ç": "c",
    }))
    return text.casefold()


def amount_terms(value):
    shown = format_amount(value)
    if not shown:
        return ""
    try:
        number = float(value)
    except (TypeError, ValueError):
        return shown
    whole = f"{number:.2f}".split(".")[0]
    return " ".join((shown, shown.replace(".", ""), whole))


def one_line(value, limit):
    text = " ".join(str(value or "").split())
    return text[:limit]


def xml_text(value):
    text = "".join(ch for ch in one_line(value, 500) if ord(ch) >= 32)
    if text[:1] in "=+-@":
        text = "'" + text
    return html.escape(text, quote=False)


def xml_number(value):
    number = float(value)
    if abs(number - round(number)) < 1e-9:
        return str(int(round(number)))
    return f"{number:.2f}"


def due_on_schedule(raw_date, repeat, today_date):
    day = parse_day(raw_date)
    if day is None:
        return False
    repeat = repeat or "Yok"
    if repeat == "Yok":
        return day <= today_date
    if day > today_date:
        return False
    if repeat == "Her gün":
        return True
    if repeat == "Her hafta":
        return day.weekday() == today_date.weekday()
    if repeat == "Her ay":
        return day.day == today_date.day
    return day == today_date


def time_reached(raw_time):
    text = (raw_time or "09:00").strip() or "09:00"
    try:
        hour, minute = text.split(":")
        now = datetime.now()
        return (now.hour, now.minute) >= (int(hour), int(minute))
    except (ValueError, TypeError):
        return True


def notify(title, body):
    if sys.platform != "darwin":
        return
    title = one_line(title, 80).replace("\\", "\\\\").replace('"', '\\"')
    body = one_line(body, 180).replace("\\", "\\\\").replace('"', '\\"')
    script = f'display notification "{body}" with title "{title}"'
    subprocess.Popen(
        ["osascript", "-e", script],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )


def open_path(path):
    if sys.platform == "win32":
        os.startfile(path)
        return
    command = ["open", str(path)] if sys.platform == "darwin" else ["xdg-open", str(path)]
    subprocess.Popen(command)


def apply_window_icon(root):
    if sys.platform != "win32":
        return
    icon = Path(__file__).resolve().parent / "vkv.ico"
    if not icon.is_file():
        return
    try:
        root.iconbitmap(default=str(icon))
    except tk.TclError:
        return


def set_dock_icon():
    icon = Path(__file__).resolve().parent / "İş Merkezi.app" / "Contents" / "Resources" / "AppIcon.icns"
    if not icon.is_file():
        return
    try:
        ctypes.cdll.LoadLibrary("/System/Library/Frameworks/AppKit.framework/AppKit")
        objc = ctypes.cdll.LoadLibrary("/usr/lib/libobjc.A.dylib")
        objc.objc_getClass.restype = ctypes.c_void_p
        objc.objc_getClass.argtypes = [ctypes.c_char_p]
        objc.sel_registerName.restype = ctypes.c_void_p
        objc.sel_registerName.argtypes = [ctypes.c_char_p]

        def msg(obj, name, *args, argtypes=None):
            send = objc.objc_msgSend
            send.restype = ctypes.c_void_p
            send.argtypes = [ctypes.c_void_p, ctypes.c_void_p] + (argtypes or [ctypes.c_void_p] * len(args))
            return send(obj, objc.sel_registerName(name), *args)

        ns_path = msg(objc.objc_getClass(b"NSString"), b"stringWithUTF8String:", str(icon).encode(), argtypes=[ctypes.c_char_p])
        image = msg(msg(objc.objc_getClass(b"NSImage"), b"alloc"), b"initWithContentsOfFile:", ns_path)
        app = msg(objc.objc_getClass(b"NSApplication"), b"sharedApplication")
        if image and app:
            msg(app, b"setApplicationIconImage:", image)
    except (OSError, AttributeError):
        return


_DROP_LIB = None


def _objc():
    ctypes.cdll.LoadLibrary("/System/Library/Frameworks/AppKit.framework/AppKit")
    objc = ctypes.cdll.LoadLibrary("/usr/lib/libobjc.A.dylib")
    objc.objc_getClass.restype = ctypes.c_void_p
    objc.objc_getClass.argtypes = [ctypes.c_char_p]
    objc.sel_registerName.restype = ctypes.c_void_p
    objc.sel_registerName.argtypes = [ctypes.c_char_p]
    objc.object_getClassName.restype = ctypes.c_char_p
    objc.object_getClassName.argtypes = [ctypes.c_void_p]
    objc.class_addMethod.restype = ctypes.c_bool
    objc.class_addMethod.argtypes = [ctypes.c_void_p, ctypes.c_void_p, ctypes.c_void_p, ctypes.c_char_p]
    return objc


def _objc_msg(objc, result, obj, name, *args, argtypes=()):
    send = objc.objc_msgSend
    send.restype = result
    send.argtypes = [ctypes.c_void_p, ctypes.c_void_p, *argtypes]
    return send(obj, objc.sel_registerName(name), *args)


def _ns_string(objc, text):
    encoded = text.encode() if isinstance(text, str) else text
    return _objc_msg(objc, ctypes.c_void_p, objc.objc_getClass(b"NSString"), b"stringWithUTF8String:", encoded, argtypes=(ctypes.c_char_p,))


def _ns_text(objc, value):
    if not value:
        return ""
    raw = _objc_msg(objc, ctypes.c_char_p, value, b"UTF8String")
    return raw.decode() if raw else ""


def _drop_lib():
    global _DROP_LIB
    if _DROP_LIB is not None or sys.platform != "darwin":
        return _DROP_LIB
    path = DATA_DIR / "araclar" / "ledger-drop.dylib"
    if not path.is_file():
        return None
    lib = ctypes.CDLL(str(path))
    lib.ledger_install.argtypes = []
    lib.ledger_watch.argtypes = [ctypes.c_void_p]
    lib.ledger_unwatch.argtypes = [ctypes.c_void_p]
    lib.ledger_hover.restype = ctypes.c_int
    lib.ledger_take.argtypes = [ctypes.c_char_p, ctypes.c_int]
    lib.ledger_take.restype = ctypes.c_int
    lib.ledger_install()
    _DROP_LIB = lib
    return lib


def _content_view(toplevel):
    objc = _objc()
    wanted = toplevel.title().encode()
    app = _objc_msg(objc, ctypes.c_void_p, objc.objc_getClass(b"NSApplication"), b"sharedApplication")
    windows = _objc_msg(objc, ctypes.c_void_p, app, b"windows")
    count = _objc_msg(objc, ctypes.c_uint64, windows, b"count") if windows else 0
    for index in range(int(count)):
        window = _objc_msg(objc, ctypes.c_void_p, windows, b"objectAtIndex:", index, argtypes=(ctypes.c_uint64,))
        title = _objc_msg(objc, ctypes.c_void_p, window, b"title")
        if _objc_msg(objc, ctypes.c_char_p, title, b"UTF8String") != wanted:
            continue
        return _objc_msg(objc, ctypes.c_void_p, window, b"contentView")
    return None


def enable_mac_file_drop(toplevel):
    if sys.platform != "darwin":
        return False
    try:
        toplevel.update_idletasks()
        view = _content_view(toplevel)
        if not view:
            return False
        lib = _drop_lib()
        if lib is None:
            return None
        objc = _objc()
        kinds = _objc_msg(objc, ctypes.c_void_p, objc.objc_getClass(b"NSMutableArray"), b"array")
        for kind in ("public.file-url", "NSFilenamesPboardType"):
            _objc_msg(objc, None, kinds, b"addObject:", _ns_string(objc, kind), argtypes=(ctypes.c_void_p,))
        _objc_msg(objc, None, view, b"registerForDraggedTypes:", kinds, argtypes=(ctypes.c_void_p,))
        lib.ledger_watch(ctypes.c_void_p(view))
        return view
    except Exception:
        return None


def clear_mac_file_drop(widget):
    view = getattr(widget, "_drop_view", None)
    if view and _DROP_LIB is not None:
        _DROP_LIB.ledger_unwatch(ctypes.c_void_p(view))
    widget._drop_view = None


def parse_stamp(value, utc=False):
    text = str(value or "").strip()
    if not text:
        return None
    parsed = None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S"):
        try:
            parsed = datetime.strptime(text[:19], fmt)
            break
        except ValueError:
            continue
    if parsed is None:
        day = parse_day(text[:10])
        return datetime(day.year, day.month, day.day) if day else None
    if utc:
        return parsed.replace(tzinfo=timezone.utc).astimezone().replace(tzinfo=None)
    return parsed


def project_key(name, taken):
    letters = "".join(ch for ch in str(name or "").upper() if ch.isalnum())
    base = (letters[:4] or "PRJ")[:4]
    candidate = base
    number = 2
    while candidate in taken:
        candidate = f"{base[:3]}{number}"
        number += 1
    return candidate


def task_waiting_since(row):
    if "waiting_since" in row.keys() and row["waiting_since"] is not None:
        return parse_stamp(row["waiting_since"]) if row["waiting_since"] else None
    if "created_at" in row.keys() and row["created_at"]:
        return parse_stamp(row["created_at"], utc=True)
    return None


def task_sla_moment(row):
    if row["status"] in ("Tamamlandı", "İptal"):
        return None
    start = task_waiting_since(row)
    if start is None:
        return None
    due = start + timedelta(hours=24)
    scheduled = parse_stamp(row["sla_next"]) if "sla_next" in row.keys() and row["sla_next"] else None
    return scheduled or due


def stamp_task_wait(data, previous):
    status = data.get("status") or "Bekliyor"
    now = datetime.now().replace(microsecond=0)
    if status in ("Tamamlandı", "İptal"):
        data["waiting_since"] = ""
        data["sla_next"] = ""
        return
    if status != "Bekliyor":
        start = task_waiting_since(previous) if previous is not None else None
        if start is not None and now >= start + timedelta(hours=24):
            data["waiting_since"] = start.strftime("%Y-%m-%d %H:%M:%S")
            data["sla_next"] = previous["sla_next"] or (start + timedelta(hours=24)).strftime("%Y-%m-%d %H:%M:%S")
            return
        data["waiting_since"] = ""
        data["sla_next"] = ""
        return
    if previous is not None and (previous["status"] or "Bekliyor") == "Bekliyor":
        if previous["waiting_since"]:
            data["waiting_since"] = previous["waiting_since"]
        if previous["sla_next"]:
            data["sla_next"] = previous["sla_next"]
        return
    data["waiting_since"] = now.strftime("%Y-%m-%d %H:%M:%S")
    data["sla_next"] = (now + timedelta(hours=24)).strftime("%Y-%m-%d %H:%M:%S")


def play_alert():
    if sys.platform == "win32":
        import winsound
        winsound.MessageBeep(winsound.MB_ICONEXCLAMATION)
        return
    sound = Path("/System/Library/Sounds/Glass.aiff")
    command = ["afplay", str(sound)] if sound.is_file() else ["osascript", "-e", "beep"]
    subprocess.Popen(command, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


def try_single_instance(handle):
    if sys.platform == "win32":
        import msvcrt
        handle.seek(0, os.SEEK_END)
        if handle.tell() == 0:
            handle.write(b" ")
            handle.flush()
        handle.seek(0)
        try:
            msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
            return True
        except OSError:
            return False
    import fcntl
    try:
        fcntl.flock(handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
        return True
    except OSError:
        return False


def default_watch_dir():
    candidate = Path.home() / "Desktop" / "Faturalar"
    return candidate if candidate.is_dir() else DATA_DIR / "GelenFaturalar"


def is_within(path, root):
    try:
        path.resolve().relative_to(root.resolve())
        return True
    except (ValueError, OSError):
        return False


def acceptable_watch_dir(path):
    try:
        resolved = Path(path).expanduser().resolve()
    except OSError:
        return False
    home = Path.home().resolve()
    blocked = {
        home,
        home / "Desktop",
        home / "Documents",
        home / "Downloads",
        home / "Library",
        Path("/"),
        Path("/Users"),
    }
    if not resolved.is_dir() or resolved in blocked:
        return False
    return is_within(resolved, home)


def safe_name(name):
    base = Path(str(name or "")).name.replace("\x00", "")
    cleaned = re.sub(r"[^A-Za-z0-9._\- \u00C0-\u024F]", "_", base).strip("._ ")
    cleaned = cleaned or "ek"
    stem, suffix = os.path.splitext(cleaned)
    suffix = suffix[:8]
    if len(stem) > 80:
        stem = stem[:80]
    return f"{stem}{suffix}"


def signature_ok(ext, header):
    if header.startswith((b"MZ", b"\x7fELF", b"#!")):
        return False
    ext = ext.lower()
    if ext == ".pdf":
        return header.lstrip().startswith(b"%PDF")
    if ext == ".png":
        return header.startswith(b"\x89PNG")
    if ext in {".jpg", ".jpeg"}:
        return header.startswith(b"\xff\xd8\xff")
    if ext == ".webp":
        return header.startswith(b"RIFF") and header[8:12] == b"WEBP"
    if ext in {".tif", ".tiff"}:
        return header[:2] in {b"II", b"MM"}
    if ext == ".heic":
        return b"ftyp" in header[:16]
    return False


def file_signature_ok(path):
    ext = path.suffix.lower()
    if ext not in ALLOWED_EXT:
        return False
    try:
        with path.open("rb") as handle:
            return signature_ok(ext, handle.read(16))
    except OSError:
        return False


def assert_openable(path, roots):
    candidate = Path(path)
    if not candidate.is_file() or candidate.is_symlink():
        raise ValueError("Dosya açılamıyor.")
    resolved = candidate.resolve()
    if resolved.suffix.lower() not in ALLOWED_EXT or not file_signature_ok(resolved):
        raise ValueError("Yalnızca doğrulanmış PDF ve görüntü açılır.")
    try:
        if resolved.stat().st_size > MAX_BYTES:
            raise ValueError("Dosya 20 MB sınırını aşıyor.")
    except OSError as exc:
        raise ValueError("Dosya okunamadı.") from exc
    if not any(is_within(resolved, root) for root in roots if root is not None):
        raise ValueError("Dosya izin verilen klasörün dışında.")
    return resolved


LEDGER_PRESETS = {
    "Gelir": ["Satış", "Hizmet", "Tahsilat", "Kira geliri", "Faiz", "Diğer gelir"],
    "Gider": ["Kira", "Maaş", "Fatura", "Vergi", "Market", "Ulaşım", "Ofis", "Bakım", "Diğer gider"],
}


def _ledger_hidden(db, kind):
    chosen = kind if kind in LEDGER_PRESETS else "Gider"
    try:
        hidden = json.loads(db.get_setting("ledger_hidden_categories", "") or "{}")
    except (json.JSONDecodeError, TypeError, ValueError):
        hidden = {}
    if not isinstance(hidden, dict):
        return chosen, set(), {}
    blocked = {fold(item) for item in hidden.get(chosen, []) if isinstance(item, str)}
    return chosen, blocked, hidden


def ledger_category_names(db, kind):
    chosen, blocked, _hidden = _ledger_hidden(db, kind)
    names = []
    seen = set()

    def add(name):
        text = " ".join(str(name or "").split())[:40]
        key = fold(text)
        if not text or key in seen or key in blocked:
            return
        seen.add(key)
        names.append(text)

    for name in LEDGER_PRESETS[chosen]:
        add(name)
    try:
        saved = json.loads(db.get_setting("ledger_categories", "") or "{}")
    except (json.JSONDecodeError, TypeError, ValueError):
        saved = {}
    if isinstance(saved, dict):
        for name in saved.get(chosen, []):
            add(name)
    for row in db.all("transactions"):
        if (row["kind"] or "") == chosen:
            add(row["category"])
    return names


def remember_ledger_category(db, kind, name):
    chosen, _blocked, hidden = _ledger_hidden(db, kind)
    text = " ".join(str(name or "").split())[:40]
    if not text:
        return
    if isinstance(hidden, dict):
        hidden[chosen] = [item for item in hidden.get(chosen, []) if isinstance(item, str) and fold(item) != fold(text)]
        db.set_setting("ledger_hidden_categories", json.dumps(hidden, ensure_ascii=False))
    if fold(text) in {fold(item) for item in LEDGER_PRESETS[chosen]}:
        return
    try:
        saved = json.loads(db.get_setting("ledger_categories", "") or "{}")
    except (json.JSONDecodeError, TypeError, ValueError):
        saved = {}
    if not isinstance(saved, dict):
        saved = {}
    current = [item for item in saved.get(chosen, []) if isinstance(item, str)]
    if fold(text) not in {fold(item) for item in current}:
        current.append(text)
    saved[chosen] = current[:40]
    db.set_setting("ledger_categories", json.dumps(saved, ensure_ascii=False))


def forget_ledger_category(db, kind, name):
    chosen, _blocked, hidden = _ledger_hidden(db, kind)
    text = " ".join(str(name or "").split())[:40]
    if not text:
        return
    try:
        saved = json.loads(db.get_setting("ledger_categories", "") or "{}")
    except (json.JSONDecodeError, TypeError, ValueError):
        saved = {}
    if not isinstance(saved, dict):
        saved = {}
    saved[chosen] = [item for item in saved.get(chosen, []) if isinstance(item, str) and fold(item) != fold(text)][:40]
    db.set_setting("ledger_categories", json.dumps(saved, ensure_ascii=False))
    if not isinstance(hidden, dict):
        hidden = {}
    blocked = [item for item in hidden.get(chosen, []) if isinstance(item, str) and fold(item) != fold(text)]
    blocked.append(text)
    hidden[chosen] = blocked[:80]
    db.set_setting("ledger_hidden_categories", json.dumps(hidden, ensure_ascii=False))


def ledger_documents(raw):
    try:
        data = json.loads(raw or "[]")
    except (TypeError, json.JSONDecodeError):
        return []
    if not isinstance(data, list):
        return []
    documents = []
    for item in data[:30]:
        if not isinstance(item, dict):
            continue
        path = str(item.get("path") or "")
        name = str(item.get("name") or Path(path).name)[:120]
        if path and name:
            documents.append({"path": path, "name": name})
    return documents


def store_ledger_file(src):
    source = Path(src)
    if not source.is_file() or source.is_symlink():
        raise ValueError("Dosya kullanılamıyor.")
    resolved = source.resolve()
    if resolved.suffix.lower() not in ALLOWED_EXT or not file_signature_ok(resolved):
        raise ValueError("Yalnızca doğrulanmış PDF ve görüntü eklenebilir.")
    try:
        size = resolved.stat().st_size
    except OSError as exc:
        raise ValueError("Dosya okunamadı.") from exc
    if size <= 0 or size > MAX_BYTES:
        raise ValueError("Dosya boş ya da 20 MB sınırının üstünde.")
    LEDGER_DIR.mkdir(parents=True, exist_ok=True)
    os.chmod(LEDGER_DIR, 0o700)
    dest = LEDGER_DIR / f"{datetime.now().strftime('%Y%m%d-%H%M%S')}-{secrets.token_hex(3)}-{safe_name(resolved.name)}"
    shutil.copy2(resolved, dest)
    os.chmod(dest, 0o600)
    if not is_within(dest, LEDGER_DIR):
        dest.unlink(missing_ok=True)
        raise ValueError("Dosya kaydedilemedi.")
    return str(dest)


def remove_ledger_file(path):
    if not path:
        return
    candidate = Path(path)
    try:
        resolved = candidate.resolve()
    except OSError:
        return
    if candidate.is_symlink() or not resolved.is_file() or not is_within(resolved, LEDGER_DIR):
        return
    try:
        resolved.unlink()
    except OSError:
        pass


def store_attachment(src, watch_root):
    if not src:
        return ""
    source = Path(src)
    if not source.is_file() or source.is_symlink():
        raise ValueError("Dosya kullanılamıyor.")
    resolved = assert_openable(source, [ATTACH_DIR, watch_root, source.parent])
    if is_within(resolved, ATTACH_DIR) or (watch_root and is_within(resolved, watch_root)):
        return str(resolved)
    ATTACH_DIR.mkdir(parents=True, exist_ok=True)
    os.chmod(ATTACH_DIR, 0o700)
    dest = ATTACH_DIR / f"{datetime.now().strftime('%Y%m%d-%H%M%S')}-{secrets.token_hex(3)}-{safe_name(resolved.name)}"
    shutil.copy2(resolved, dest)
    os.chmod(dest, 0o600)
    return str(dest)


REPORT_EXT = {".pdf", ".png", ".jpg", ".jpeg", ".webp", ".tif", ".tiff", ".heic", ".eml"}


def report_signature_ok(path):
    ext = path.suffix.lower()
    if ext not in REPORT_EXT:
        return False
    try:
        with path.open("rb") as handle:
            header = handle.read(512)
    except OSError:
        return False
    if header.startswith((b"MZ", b"\x7fELF", b"#!", b"\x00")):
        return False
    if ext == ".eml":
        text = header.decode("utf-8", "ignore")
        return any(token in text for token in ("From:", "From ", "Received:", "MIME-Version:", "Subject:", "Return-Path:", "Date:"))
    return signature_ok(ext, header[:16])


OFFER_EXT = {".pdf", ".eml", ".png", ".jpg", ".jpeg", ".webp", ".tif", ".tiff", ".heic", ".docx", ".xlsx"}


def offer_kind(path):
    ext = Path(path or "").suffix.lower()
    return {
        ".pdf": "PDF",
        ".eml": "E-posta",
        ".docx": "Word",
        ".xlsx": "Excel",
        ".png": "Görüntü",
        ".jpg": "Görüntü",
        ".jpeg": "Görüntü",
        ".webp": "Görüntü",
        ".tif": "Görüntü",
        ".tiff": "Görüntü",
        ".heic": "Görüntü",
    }.get(ext, "Dosya")


def offer_signature_ok(path):
    ext = path.suffix.lower()
    if ext not in OFFER_EXT:
        return False
    try:
        with path.open("rb") as handle:
            header = handle.read(512)
    except OSError:
        return False
    if not header or header.startswith((b"MZ", b"\x7fELF", b"#!", b"\x00")):
        return False
    if ext == ".eml":
        text = header.decode("utf-8", "ignore")
        return any(token in text for token in ("From:", "From ", "Received:", "MIME-Version:", "Subject:", "Return-Path:", "Date:"))
    if ext in {".docx", ".xlsx"}:
        return header.startswith(b"PK\x03\x04")
    return signature_ok(ext, header[:16])


def assert_offer_file(path):
    candidate = Path(path)
    if not candidate.is_file() or candidate.is_symlink():
        raise ValueError("Dosya açılamıyor.")
    resolved = candidate.resolve()
    if not offer_signature_ok(resolved):
        raise ValueError("Yalnızca PDF, kayıtlı e-posta (.eml), görüntü, Word (.docx) veya Excel (.xlsx) eklenebilir.")
    try:
        size = resolved.stat().st_size
    except OSError as exc:
        raise ValueError("Dosya okunamadı.") from exc
    if size <= 0 or size > MAX_BYTES:
        raise ValueError("Dosya boş ya da 20 MB sınırının üstünde.")
    return resolved


def store_offer_file(src):
    resolved = assert_offer_file(src)
    if is_within(resolved, OFFER_DIR):
        return str(resolved)
    OFFER_DIR.mkdir(parents=True, exist_ok=True)
    os.chmod(OFFER_DIR, 0o700)
    dest = OFFER_DIR / f"{datetime.now().strftime('%Y%m%d-%H%M%S')}-{secrets.token_hex(3)}-{safe_name(resolved.name)}"
    shutil.copy2(resolved, dest)
    os.chmod(dest, 0o600)
    return str(dest)


def remove_offer_file(path):
    if not path:
        return
    candidate = Path(path)
    try:
        resolved = candidate.resolve()
    except OSError:
        return
    if candidate.is_symlink() or not resolved.is_file() or not is_within(resolved, OFFER_DIR):
        return
    try:
        resolved.unlink()
    except OSError:
        pass


def openable_offer(path):
    resolved = assert_offer_file(path)
    if not is_within(resolved, OFFER_DIR):
        raise ValueError("Dosya izin verilen klasörün dışında.")
    return resolved


OFFER_FEATURES = (
    ("kurulum", "Kurulum"),
    ("montaj", "Montaj"),
    ("egitim", "Eğitim"),
    ("nakliye", "Nakliye"),
    ("bakim", "Bakım"),
    ("devreye alma", "Devreye alma"),
    ("yedek parca", "Yedek parça"),
    ("lisans", "Lisans"),
    ("destek", "Destek"),
)


def offer_tool():
    if sys.platform != "darwin":
        raise ValueError("PDF okuma aracı yalnız macOS üzerinde hazır.")
    with _REPORT_LOCK:
        source = Path(__file__).resolve().parent / "İş Merkezi.app" / "Contents" / "Resources" / "teklif-oku.swift"
        if not source.is_file():
            raise ValueError("Teklif okuma aracı eksik.")
        tool_dir = DATA_DIR / "araclar"
        tool = tool_dir / "teklif-oku"
        stamp = str(source.stat().st_mtime_ns)
        marker = tool_dir / "teklif-oku.stamp"
        if tool.is_file() and marker.is_file() and marker.read_text(encoding="utf-8") == stamp:
            return tool
        tool_dir.mkdir(parents=True, exist_ok=True)
        os.chmod(tool_dir, 0o700)
        result = subprocess.run(["swiftc", "-O", "-o", str(tool), str(source)], capture_output=True, text=True)
        if result.returncode != 0 or not tool.is_file():
            raise ValueError(result.stderr.strip() or "Teklif okuma aracı derlenemedi.")
        marker.write_text(stamp, encoding="utf-8")
        os.chmod(tool, 0o700)
        return tool


def _tool_text(path):
    result = subprocess.run([str(offer_tool()), str(path)], capture_output=True, timeout=60)
    return result.stdout.decode("utf-8", "ignore")[:20000]


def _plain_html(value):
    text = re.sub(r"(?is)<(script|style).*?>.*?</\1>", " ", value or "")
    text = re.sub(r"<[^>]+>", " ", text)
    return html.unescape(text)


def _eml_text(path):
    message = email.message_from_bytes(path.read_bytes()[:MAX_BYTES], policy=email.policy.default)
    chunks = []
    subject = str(message.get("Subject") or "").strip()
    if subject:
        chunks.append(subject)
    folder = DATA_DIR / "araclar"
    for part in message.walk():
        if part.get_content_disposition() == "attachment" and part.get_content_type() != "application/pdf":
            continue
        kind = part.get_content_type()
        try:
            if kind == "text/plain":
                chunks.append(str(part.get_content()))
            elif kind == "text/html":
                chunks.append(_plain_html(str(part.get_content())))
            elif kind == "application/pdf":
                payload = part.get_payload(decode=True) or b""
                if not payload.lstrip().startswith(b"%PDF") or len(payload) > MAX_BYTES:
                    continue
                folder.mkdir(parents=True, exist_ok=True)
                os.chmod(folder, 0o700)
                temp = folder / f"oku-{secrets.token_hex(4)}.pdf"
                temp.write_bytes(payload)
                os.chmod(temp, 0o600)
                try:
                    chunks.append(_tool_text(temp))
                finally:
                    temp.unlink(missing_ok=True)
        except (email.errors.MessageError, subprocess.SubprocessError, OSError, ValueError, KeyError, TypeError):
            continue
    return "\n".join(chunks)[:20000]


def _office_text(path, names):
    chunks = []
    with zipfile.ZipFile(path) as archive:
        for name in names:
            if name not in archive.namelist():
                continue
            info = archive.getinfo(name)
            if info.file_size > 2 * 1024 * 1024:
                continue
            root = ET.fromstring(archive.read(name))
            for node in root.iter():
                if node.text and (node.tag.endswith("}t") or node.tag.endswith("}v")):
                    chunks.append(node.text)
                if sum(len(item) for item in chunks) > 20000:
                    return "\n".join(chunks)[:20000]
    return "\n".join(chunks)[:20000]


def read_offer_text(path):
    resolved = assert_offer_file(path)
    ext = resolved.suffix.lower()
    try:
        if ext == ".eml":
            return _eml_text(resolved)
        if ext == ".docx":
            return _office_text(resolved, ["word/document.xml"])
        if ext == ".xlsx":
            return _office_text(resolved, ["xl/sharedStrings.xml", "xl/worksheets/sheet1.xml"])
        if ext in {".pdf", ".png", ".jpg", ".jpeg", ".webp", ".tif", ".tiff", ".heic"}:
            return _tool_text(resolved)
    except (OSError, ValueError, subprocess.SubprocessError, zipfile.BadZipFile, ET.ParseError, email.errors.MessageError):
        return ""
    return ""


OFFER_URL_BYTES = 3 * 1024 * 1024
_URL_AGENT = "Mozilla/5.0 (Macintosh; Intel Mac OS X 14_0) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Safari/605.1.15"
_SLUG_SCREEN = r"(?=\d{2,3}hz\b|tasinabilir|laptop|notebook|inc\b|inch\b|fhd|wuxga|qhd|oled|ips|led\b|uhd|4k\b|smart|televizyon|monitor)"


def clean_offer_url(value):
    text = str(value or "").strip()
    if not text:
        return ""
    if not re.match(r"(?i)^https?://", text):
        text = "https://" + text
    parts = urllib.parse.urlsplit(text)
    if parts.scheme.lower() not in {"http", "https"} or not parts.hostname or "." not in parts.hostname:
        raise ValueError("Adres http veya https ile başlayan bir internet sayfası olmalı.")
    if len(text) > 2000 or any(ch.isspace() for ch in text):
        raise ValueError("Adres geçersiz.")
    return text


def offer_url_host(url):
    try:
        host = urllib.parse.urlsplit(url).hostname or ""
    except ValueError:
        return ""
    return host[4:] if host.startswith("www.") else host


def offer_url_slug_text(url):
    try:
        path = urllib.parse.unquote(urllib.parse.urlsplit(url).path)
    except ValueError:
        return ""
    words = fold(re.sub(r"[-_/+]+", " ", path))
    words = re.sub(r"\b(?:p|pm|dp|pd)\s+(?=[a-z0-9]*\d)[a-z0-9]{6,}\s*$", "", words.strip())
    words = re.sub(rf"(?<![\d.])(1\d|[2-9]\d)\s(\d)\s{_SLUG_SCREEN}", r'\1.\2" ', words)
    words = re.sub(rf"(?<![\d.\"])(1[0-8]|[2-9]\d)\s{_SLUG_SCREEN}", r'\1" ', words)
    words = re.sub(r"\.html?$", "", words.strip())
    return " ".join(words.split())


def _ld_items(body):
    items = []
    for raw in re.findall(r"(?is)<script[^>]*application/ld\+json[^>]*>(.*?)</script>", body)[:12]:
        try:
            data = json.loads(html.unescape(raw.strip()))
        except (json.JSONDecodeError, ValueError):
            continue
        stack = [data]
        while stack and len(items) < 40:
            node = stack.pop()
            if isinstance(node, list):
                stack.extend(node)
            elif isinstance(node, dict):
                items.append(node)
                for key in ("@graph", "itemListElement", "mainEntity"):
                    if key in node:
                        stack.append(node[key])
    return items


def _meta(body, name):
    for pattern in (
        rf'(?is)<meta[^>]+(?:property|name|itemprop)=["\']{re.escape(name)}["\'][^>]+content=["\']([^"\']*)["\']',
        rf'(?is)<meta[^>]+content=["\']([^"\']*)["\'][^>]+(?:property|name|itemprop)=["\']{re.escape(name)}["\']',
    ):
        match = re.search(pattern, body)
        if match:
            return " ".join(html.unescape(match.group(1)).split())
    return ""


def _page_price(items, body):
    for node in items:
        if "product" not in fold(str(node.get("@type", ""))):
            continue
        offers = node.get("offers")
        for offer in offers if isinstance(offers, list) else [offers]:
            if not isinstance(offer, dict):
                continue
            for key in ("price", "lowPrice"):
                value = offer.get(key)
                if value in (None, ""):
                    continue
                try:
                    return round(float(str(value).replace(",", ".")), 2)
                except ValueError:
                    continue
    for name in ("product:price:amount", "og:price:amount", "price"):
        value = _meta(body, name)
        if value:
            try:
                return parse_amount(value, False)
            except ValueError:
                continue
    return None


def _no_ftp_redirect():
    class Redirect(urllib.request.HTTPRedirectHandler):
        def redirect_request(self, req, fp, code, msg, headers, newurl):
            if urllib.parse.urlsplit(newurl).scheme.lower() not in {"http", "https"}:
                return None
            return super().redirect_request(req, fp, code, msg, headers, newurl)

    return urllib.request.build_opener(Redirect)


_PAGE_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
    "Accept-Language": "tr-TR,tr;q=0.9,en-US;q=0.8",
    "Sec-Fetch-Dest": "document",
    "Sec-Fetch-Mode": "navigate",
    "Sec-Fetch-Site": "none",
    "Sec-Fetch-User": "?1",
    "Upgrade-Insecure-Requests": "1",
    "sec-ch-ua": '"Chromium";v="131", "Not_A Brand";v="24"',
    "sec-ch-ua-mobile": "?0",
    "sec-ch-ua-platform": '"macOS"',
}
_PAGE_LOCK = threading.Lock()


def page_tool():
    if sys.platform != "darwin":
        raise ValueError("Tarayıcıyla okuma yalnız macOS üzerinde hazır.")
    with _PAGE_LOCK:
        source = Path(__file__).resolve().parent / "İş Merkezi.app" / "Contents" / "Resources" / "sayfa-oku.swift"
        if not source.is_file():
            raise ValueError("Sayfa okuma aracı eksik.")
        tool_dir = DATA_DIR / "araclar"
        tool = tool_dir / "sayfa-oku"
        stamp = str(source.stat().st_mtime_ns)
        marker = tool_dir / "sayfa-oku.stamp"
        if tool.is_file() and marker.is_file() and marker.read_text(encoding="utf-8") == stamp:
            return tool
        tool_dir.mkdir(parents=True, exist_ok=True)
        os.chmod(tool_dir, 0o700)
        result = subprocess.run(["swiftc", "-O", "-o", str(tool), str(source)], capture_output=True, text=True)
        if result.returncode != 0 or not tool.is_file():
            raise ValueError(result.stderr.strip() or "Sayfa okuma aracı derlenemedi.")
        marker.write_text(stamp, encoding="utf-8")
        os.chmod(tool, 0o700)
        return tool


def _plain_get(url):
    request = urllib.request.Request(url, headers=_PAGE_HEADERS)
    with _no_ftp_redirect().open(request, timeout=12) as response:
        kind = response.headers.get("Content-Type", "")
        if "html" not in kind.lower() and "xml" not in kind.lower():
            raise ValueError("Sayfa HTML değil.")
        charset = response.headers.get_content_charset() or "utf-8"
        return response.read(OFFER_URL_BYTES).decode(charset, "ignore")


def _browser_get(url):
    result = subprocess.run([str(page_tool()), url], capture_output=True, timeout=45)
    try:
        data = json.loads(result.stdout.decode("utf-8", "ignore") or "{}")
    except json.JSONDecodeError as exc:
        raise ValueError("Tarayıcı sayfayı okuyamadı.") from exc
    if not isinstance(data, dict) or data.get("error"):
        raise ValueError(str(data.get("error") if isinstance(data, dict) else "") or "Tarayıcı sayfayı okuyamadı.")
    nodes = []
    for raw in data.get("ld") or []:
        try:
            parsed = json.loads(str(raw))
        except (json.JSONDecodeError, ValueError):
            continue
        stack = [parsed]
        while stack and len(nodes) < 40:
            node = stack.pop()
            if isinstance(node, list):
                stack.extend(node)
            elif isinstance(node, dict):
                nodes.append(node)
                for key in ("@graph", "itemListElement", "mainEntity"):
                    if key in node:
                        stack.append(node[key])
    def listed(key):
        return [tuple(item[:2]) for item in data.get(key) or [] if isinstance(item, list) and len(item) >= 2]

    return {
        "title": str(data.get("ogTitle") or data.get("title") or ""),
        "description": str(data.get("description") or ""),
        "price_meta": str(data.get("price") or ""),
        "nodes": nodes,
        "visible": str(data.get("text") or ""),
        "specs": _merge_specs(_ld_spec_pairs(nodes) + listed("specs"), listed("jsonSpecs")),
    }


def _html_page(body):
    title_tag = re.search(r"(?is)<title[^>]*>(.*?)</title>", body)
    nodes = _ld_items(body)
    return {
        "title": _meta(body, "og:title") or (" ".join(html.unescape(title_tag.group(1)).split()) if title_tag else ""),
        "description": _meta(body, "og:description") or _meta(body, "description"),
        "price_meta": _meta(body, "product:price:amount") or _meta(body, "og:price:amount"),
        "nodes": nodes,
        "visible": _plain_html(body),
        "specs": _merge_specs(_ld_spec_pairs(nodes) + _table_spec_pairs(body), _json_spec_pairs(body)),
    }


_ADS = re.compile(r"\bReklam\b|\bSponsorlu\b|Markaların Seçtikleri|\bSponsored\b", re.I)
_OTHER_PRODUCTS = re.compile(
    _ADS.pattern + r"|Benzer Ürünler|Bunlar da İlgin|Sizin [İi]çin Seçt|Son Gez|Birlikte Alın|"
    r"Bu Ürün(?:e|ü) Bakanlar|Önerilen Ürünler|İlgini Çekebilecek|Çok Satanlar|Customers who (?:viewed|bought)",
    re.I,
)


def _product_region(visible, title):
    flat = " ".join(str(visible or "").split())
    start = 0
    if title:
        spot = fold(flat).find(fold(title[:40]))
        if spot >= 0:
            start = spot
    end = (_OTHER_PRODUCTS if start else _ADS).search(flat, start + len(title or ""))
    return flat[start:end.start() if end else None], start > 0


def _text_price(region, anchored):
    for chunk in (region[:2500],) if anchored else (region[:6000],):
        for match in re.finditer(r"(?<![\d.,])(\d{1,3}(?:\.\d{3})+,\d{2}|\d{1,6},\d{2})\s*(?:TL|₺)", chunk):
            if re.search(r"(?i)(\d\s*x|taksit|kazan\w*|kargo\w*)\W*$", chunk[max(0, match.start() - 30):match.start()]):
                continue
            try:
                value = parse_amount(match.group(1), False)
            except ValueError:
                continue
            if value and value > 0:
                return value
    return None


SPEC_HEADER = "Teknik özellikler:"
_SPEC_NOISE_LABEL = re.compile(r"^(?:[a-z]+[A-Z]|[A-Z]+[a-z]+[A-Z]|Is[A-Z])|^[\d\W_]+$|^(?i:stok|barkod|sku)\b")
_SPEC_NOISE_VALUE = re.compile(r"(?i)^(?:https?:|www\.|/|#[0-9a-f]{3,8}$|true$|false$|null$|none$|undefined$|linear-gradient|rgba?\(|detaylar\w* gör)")
_SPEC_PAIR = re.compile(
    r'"(?:name|key|label|attributeName)"\s*:\s*"((?:[^"\\]|\\.){2,60})"\s*,\s*'
    r'"(?:property|value|unitText|attributeValue|values?)"\s*:\s*"((?:[^"\\]|\\.){1,200})"'
)


def clean_spec_pairs(pairs, limit=80):
    clean, seen = [], set()
    for label, value in pairs:
        label = " ".join(str(label or "").split()).strip(" :")
        value = " ".join(str(value or "").split()).strip()
        if not label or not value or len(label) > 60 or len(value) > 200 or fold(label) == fold(value):
            continue
        if _SPEC_NOISE_LABEL.search(label) or _SPEC_NOISE_VALUE.search(value) or re.fullmatch(r"[\d.,]+\s*\(\d+\)", value):
            continue
        key = fold(label)
        if key in seen:
            continue
        seen.add(key)
        clean.append((label, value))
        if len(clean) >= limit:
            break
    return clean


def _json_text(raw):
    try:
        return json.loads(f'"{raw}"')
    except (json.JSONDecodeError, ValueError):
        return raw


def _merge_specs(primary, loose):
    specs = clean_spec_pairs(primary)
    if len(specs) < 5:
        return clean_spec_pairs(specs + list(loose))
    core = [pair for pair in loose if re.search(r"\b(?:ram|bellek|ssd|depolama|disk|islemci|ekran karti)\b", fold(pair[0]))]
    return clean_spec_pairs(specs + core)


def _json_spec_pairs(body):
    return [(html.unescape(_json_text(match.group(1))), html.unescape(_json_text(match.group(2)))) for match in _SPEC_PAIR.finditer(body or "")]


def _table_spec_pairs(body):
    pairs = []
    for row in re.findall(r"(?is)<tr[^>]*>(.*?)</tr>", body or "")[:300]:
        cells = re.findall(r"(?is)<t[hd][^>]*>(.*?)</t[hd]>", row)
        if len(cells) == 2:
            pairs.append(tuple(" ".join(_plain_html(cell).split()) for cell in cells))
    for term, detail in re.findall(r"(?is)<dt[^>]*>(.*?)</dt>\s*<dd[^>]*>(.*?)</dd>", body or "")[:200]:
        pairs.append((" ".join(_plain_html(term).split()), " ".join(_plain_html(detail).split())))
    return pairs


def _ld_spec_pairs(nodes):
    pairs = []
    for node in nodes:
        if "product" not in fold(str(node.get("@type", ""))):
            continue
        for prop in node.get("additionalProperty") or []:
            if isinstance(prop, dict) and prop.get("name"):
                value = prop.get("value")
                if value in (None, ""):
                    value = prop.get("unitText")
                if isinstance(value, list):
                    value = ", ".join(str(item) for item in value)
                pairs.append((prop.get("name"), value))
    return pairs


def _teknosa_spec_pairs(url):
    code = re.search(r"-p-(\d{5,})(?:[/?#]|$)", url)
    if not code or offer_url_host(url) != "teknosa.com":
        return []
    headers = dict(_PAGE_HEADERS)
    headers.update({
        "Accept": "application/json, text/javascript, */*; q=0.01",
        "X-Requested-With": "XMLHttpRequest",
        "Sec-Fetch-Mode": "cors",
        "Sec-Fetch-Dest": "empty",
        "Sec-Fetch-Site": "same-origin",
        "Referer": url,
    })
    query = urllib.parse.urlencode({"productCode": code.group(1)})
    request = urllib.request.Request(f"https://www.teknosa.com/getClassificationData?{query}", headers=headers)
    try:
        with _no_ftp_redirect().open(request, timeout=10) as response:
            data = json.loads(response.read(OFFER_URL_BYTES).decode("utf-8", "ignore"))
    except (urllib.error.URLError, OSError, ValueError, TimeoutError):
        return []
    pairs = []
    for group in data if isinstance(data, list) else []:
        for feature in (group or {}).get("features") or []:
            values = [str(item.get("value")) for item in feature.get("featureValues") or [] if isinstance(item, dict) and item.get("value")]
            if feature.get("name") and values:
                pairs.append((feature["name"], ", ".join(values)))
    return pairs


def offer_spec_table(text):
    pairs = []
    inside = False
    for line in str(text or "").splitlines():
        if line.strip() == SPEC_HEADER:
            inside = True
            continue
        if inside:
            if ":" not in line:
                break
            label, value = line.split(":", 1)
            pairs.append((label.strip(), value.strip()))
    return clean_spec_pairs(pairs, 150)


def _page_letters(page):
    return len(re.findall(r"[a-zA-ZçğıöşüÇĞİÖŞÜ]", page["visible"] or ""))


def fetch_offer_page(url):
    url = clean_offer_url(url)
    slug = offer_url_slug_text(url)
    result = {"url": url, "title": "", "price": None, "text": "", "blocked": False, "note": "", "specs": []}
    page = None
    problem = ""
    try:
        page = _html_page(_plain_get(url))
        if _page_letters(page) < 400:
            problem = "sayfa boş geldi"
            page = None
    except urllib.error.HTTPError as exc:
        problem = f"site {exc.code} verdi"
    except (urllib.error.URLError, OSError, ValueError, TimeoutError, EOFError) as exc:
        problem = one_line(getattr(exc, "reason", exc), 60)
    via = "Sayfa okundu."
    if page is None and sys.platform == "darwin":
        try:
            page = _browser_get(url)
            if _page_letters(page) < 200:
                problem = "tarayıcıda da içerik gelmedi"
                page = None
            else:
                via = "Sayfa tarayıcı motoruyla okundu."
        except (OSError, ValueError, subprocess.SubprocessError) as exc:
            problem = one_line(exc, 80)
            page = None
    head = []
    specs = []
    if page is not None:
        specs = clean_spec_pairs(_teknosa_spec_pairs(url) + page["specs"])
        if len(specs) < 5 and via == "Sayfa okundu." and sys.platform == "darwin":
            try:
                rendered = _browser_get(url)
                if len(rendered["specs"]) > len(specs):
                    specs = clean_spec_pairs(specs + rendered["specs"])
                    via = "Sayfa okundu; özellikler tarayıcı motoruyla alındı."
            except (OSError, ValueError, subprocess.SubprocessError):
                pass
        result["specs"] = specs
        title = re.split(r"\s+[|–-]\s+(?=[^|–-]*$)", page["title"])[0].strip() if page["title"] else ""
        result["title"] = title[:160]
        head.append(title)
        head.append(page["description"])
        about = []
        for node in page["nodes"]:
            if "product" in fold(str(node.get("@type", ""))):
                head.append(str(node.get("name") or ""))
                brand = node.get("brand")
                if isinstance(brand, dict):
                    brand = brand.get("name")
                if brand:
                    head.append(f"Marka: {brand}")
                about.append(str(node.get("description") or "")[:1500])
        if specs:
            head.append(SPEC_HEADER)
            head.extend(f"{label}: {value}" for label, value in specs)
            head.append(None)
        head.extend(about)
        price = _page_price(page["nodes"], "")
        if price is None and page["price_meta"]:
            try:
                price = parse_amount(page["price_meta"], False)
            except ValueError:
                price = None
        region, anchored = _product_region(page["visible"], title)
        if not price or price <= 0:
            price = _text_price(region, anchored)
        result["price"] = price
        result["note"] = via
    else:
        result["blocked"] = True
        result["note"] = f"Sayfa açılamadı ({problem}); ürün bilgisi adresten okundu." if problem else "Ürün bilgisi adresten okundu."
    lines = ["" if line is None else line for line in head if line is None or line]
    if slug:
        lines.append(f"Adres: {slug}")
    if page is not None:
        lines.append(region[:6000])
    result["text"] = "\n".join(lines)[:20000]
    if not result["title"] and slug:
        result["title"] = _pretty_words(re.sub(r'"', "", slug))[:160]
    return result


def _duration_months(number, unit):
    unit = fold(unit)
    if unit.startswith("yil"):
        return number * 12
    if unit.startswith("hafta"):
        return max(1, round(number * 7 / 30))
    if unit.startswith("gun"):
        return max(1, round(number / 30))
    return number


def _duration_days(number, unit):
    unit = fold(unit)
    if unit.startswith("hafta"):
        return number * 7
    if unit.startswith("ay"):
        return number * 30
    return number


def _feature_state(word, folded):
    for sentence in re.split(r"[.\n;]", folded):
        if word not in sentence:
            continue
        if re.search(rf"{word}.{{0,30}}(dahil degil|haric)", sentence) or re.search(rf"(dahil degil|haric).{{0,20}}{word}", sentence):
            return "exclude"
        if re.search(rf"{word}.{{0,30}}dahil", sentence) or re.search(rf"dahil.{{0,20}}{word}", sentence):
            return "include"
    return ""


def find_offer_total(folded):
    for label in ("genel toplam", "odenecek tutar", "toplam tutar", "teklif tutari"):
        match = re.search(rf"{label}[^0-9]{{0,24}}(\d{{1,3}}(?:\.\d{{3}})*(?:,\d{{2}})?)", folded)
        if not match:
            continue
        try:
            return parse_amount(match.group(1), False)
        except ValueError:
            continue
    return None


def extract_offer_profile(vendor, amount, text, has_file):
    folded = fold(text or "")
    profile = {
        "vendor": vendor or "Firma",
        "amount": amount,
        "kdv": "",
        "warranty_months": None,
        "delivery_days": None,
        "payment_days": None,
        "includes": [],
        "excludes": [],
        "discount": None,
        "validity_days": None,
        "stock": False,
        "has_text": len(re.findall(r"[a-z]", folded)) >= 40,
        "has_file": has_file,
    }
    if "kdv dahil" in folded:
        profile["kdv"] = "dahil"
    elif "kdv haric" in folded or "kdv dahil degil" in folded:
        profile["kdv"] = "haric"
    warranty = re.search(r"(\d{1,2})\s*(yil|ay)\s*garanti", folded) or re.search(r"garanti(?:\s*suresi)?\s*[:\-]?\s*(\d{1,2})\s*(yil|ay)", folded)
    unit_first = re.search(r"garanti\s*suresi\s*\((yil|ay)\)\s*[:\-]?\s*(\d{1,3})\b", folded)
    if warranty:
        profile["warranty_months"] = _duration_months(int(warranty.group(1)), warranty.group(2))
    elif unit_first:
        profile["warranty_months"] = _duration_months(int(unit_first.group(2)), unit_first.group(1))
    delivery = re.search(r"teslim(?:at)?(?:\s*suresi)?\s*[:\-]?\s*(\d{1,3})\s*(gun|hafta|ay)", folded) or re.search(r"(\d{1,3})\s*(gun|hafta|ay).{0,16}teslim", folded) or re.search(r"termin\s*[:\-]?\s*(\d{1,3})\s*(gun|hafta|ay)", folded)
    if delivery:
        profile["delivery_days"] = _duration_days(int(delivery.group(1)), delivery.group(2))
    payment = re.search(r"(\d{1,3})\s*gun\s*vade", folded) or re.search(r"vade\s*[:\-]?\s*(\d{1,3})\s*gun", folded)
    if payment:
        profile["payment_days"] = int(payment.group(1))
    elif re.search(r"(?<![a-z])pesin(?![a-z])", folded) and "pesin fiyat" not in folded:
        profile["payment_days"] = 0
    for word, _label in OFFER_FEATURES:
        state = _feature_state(word, folded)
        if state == "include":
            profile["includes"].append(word)
        elif state == "exclude":
            profile["excludes"].append(word)
    discount = re.search(r"(?:iskonto|indirim)\s*[:\-]?\s*%?\s*(\d{1,2})", folded) or re.search(r"%\s*(\d{1,2})\s*(?:iskonto|indirim)", folded)
    if discount:
        profile["discount"] = int(discount.group(1))
    validity = re.search(r"(\d{1,3})\s*gun\s*gecerli", folded)
    if validity:
        profile["validity_days"] = int(validity.group(1))
    if "hemen teslim" in folded or re.search(r"(?<![a-z])stokta(?![a-z])", folded):
        profile["stock"] = True
    if profile["amount"] is None:
        profile["amount"] = find_offer_total(folded)
    return profile


def _offer_money(amount):
    shown = format_amount(amount)
    return f"{shown} ₺" if shown else ""


def _named(items):
    return ", ".join(dict.fromkeys(item["vendor"] for item in items))


def narrate_offer(profile, profiles):
    pros = []
    cons = []
    others = [item for item in profiles if fold(item["vendor"]) != fold(profile["vendor"])]
    if profile["kdv"] == "dahil":
        pros.append("KDV dahil")
    elif profile["kdv"] == "haric":
        cons.append("KDV hariç")
    months = profile["warranty_months"]
    if months:
        label = f"Garanti {months // 12} yıl" if months % 12 == 0 else f"Garanti {months} ay"
        longer = [item for item in others if item["warranty_months"] and item["warranty_months"] < months]
        shorter = [item for item in others if item["warranty_months"] and item["warranty_months"] > months]
        if longer and not shorter:
            pros.append(f"{label}; {_named(longer)} firmasından uzun")
        elif shorter and not longer:
            cons.append(f"{label}; {_named(shorter)} firmasından kısa")
        else:
            pros.append(label)
    days = profile["delivery_days"]
    if days:
        faster = [item for item in others if item["delivery_days"] and item["delivery_days"] > days]
        slower = [item for item in others if item["delivery_days"] and item["delivery_days"] < days]
        if faster and not slower:
            pros.append(f"Teslim {days} gün; {_named(faster)} firmasından kısa")
        elif slower and not faster:
            cons.append(f"Teslim {days} gün; {_named(slower)} daha erken teslim ediyor")
        else:
            pros.append(f"Teslim süresi {days} gün")
    if profile["stock"]:
        pros.append("Stokta, hemen teslim görünüyor")
    payment = profile["payment_days"]
    if payment == 0:
        later = [item for item in others if item["payment_days"]]
        cons.append("Peşin ödeme isteniyor" + (f"; {_named(later)} vadeli" if later else ""))
    elif payment:
        sooner = [item for item in others if item["payment_days"] is not None and item["payment_days"] < payment]
        if sooner:
            pros.append(f"Ödeme vadesi {payment} gün; {_named(sooner)} firmasından uzun")
        else:
            pros.append(f"Ödeme vadesi {payment} gün")
    if profile["includes"]:
        pros.append("Dahil: " + ", ".join(label for word, label in OFFER_FEATURES if word in profile["includes"]))
    if profile["excludes"]:
        cons.append("Hariç: " + ", ".join(label for word, label in OFFER_FEATURES if word in profile["excludes"]))
    if profile["discount"]:
        pros.append(f"İskonto %{profile['discount']}")
    if profile["validity_days"] and profile["validity_days"] <= 10:
        cons.append(f"Teklif yalnızca {profile['validity_days']} gün geçerli")
    priced = [item for item in others if item["amount"] is not None]
    if profile["amount"] is not None and priced:
        best = min(priced, key=lambda item: item["amount"])
        if profile["amount"] < best["amount"]:
            rest = ", ".join(f"{item['vendor']} {_offer_money(item['amount'])}" for item in priced)
            pros.insert(0, f"Bu projedeki en düşük tutar. Diğer firmalar: {rest}")
        elif profile["amount"] > best["amount"]:
            cons.insert(0, f"{_offer_money(profile['amount'] - best['amount'])} daha yüksek; daha uygun olan {best['vendor']} ({_offer_money(best['amount'])})")
        else:
            pros.insert(0, f"Tutar {best['vendor']} ile aynı")
    elif profile["amount"] is not None and not others:
        pros.append(f"Tutar {_offer_money(profile['amount'])}. Başka firma eklenince fark burada ayrılır")
    signals = any((
        profile["kdv"], profile["warranty_months"], profile["delivery_days"] is not None,
        profile["payment_days"] is not None, profile["includes"], profile["excludes"],
        profile["discount"], profile["stock"],
    ))
    if profile["has_file"] and not profile["has_text"]:
        cons.append("Dosyadan okunabilir koşul çıkmadı")
    elif profile["has_text"] and not signals:
        cons.append("Dosyada garanti, teslim, KDV veya ödeme koşulu görünmüyor")
    elif not profile["has_file"]:
        cons.append("Karşılaştırılacak dosya yok")
    if not pros:
        pros.append("Dosyada fiyat dışında öne çıkan bir avantaj görünmüyor")
    return pros[:8], cons[:8]


COMPARE_MODES = {
    "Fiyat odaklı": {"price": 0.65, "hardware": 0.10, "warranty": 0.09, "delivery": 0.06, "payment": 0.06, "scope": 0.04},
    "Dengeli": {"price": 0.40, "hardware": 0.20, "warranty": 0.15, "delivery": 0.09, "payment": 0.07, "scope": 0.09},
    "Kalite odaklı": {"price": 0.20, "hardware": 0.30, "warranty": 0.22, "delivery": 0.10, "payment": 0.06, "scope": 0.12},
}
COMPARE_LABELS = {"price": "Fiyat", "hardware": "Donanım", "warranty": "Garanti", "delivery": "Teslim", "payment": "Vade", "scope": "Kapsam"}
_SPEC_UNITS = re.compile(r"(\d{1,5}(?:[.,]\d{1,2})?)\s*(inc|inch|tb|gb|ghz|mhz|hz|kw|btu|mp|nit|watt|mah|core|cekirdek)(?![a-z])")
_SPEC_UNIT_LABEL = {
    "inc": "inç", "inch": "inç", "tb": "TB", "gb": "GB", "ghz": "GHz", "mhz": "MHz", "hz": "Hz",
    "kw": "kW", "btu": "BTU", "mp": "MP", "nit": "nit", "watt": "W", "mah": "mAh", "core": "çekirdek", "cekirdek": "çekirdek",
}
_SPEC_WORDS = (
    ("8k", "8K"), ("4k", "4K"), ("uhd", "UHD"), ("full hd", "Full HD"), ("qled", "QLED"), ("oled", "OLED"),
    ("mini led", "Mini LED"), ("smart", "Smart"), ("android", "Android"), ("google tv", "Google TV"),
    ("wifi", "Wi-Fi"), ("ssd", "SSD"), ("nvme", "NVMe"), ("hdd", "HDD"), ("raid", "RAID"), ("rack", "Rack"),
    ("inverter", "Inverter"), ("yedekli guc", "Yedekli güç"),
)
_TITLE_STOP = {
    "satis", "satisi", "teklif", "teklifi", "alim", "alimi", "icin", "ve", "ile", "adet", "urun", "urunu",
    "fiyat", "fiyati", "tedarik", "tedariki", "hizmet", "hizmeti", "kurulum", "kurulumu",
}


def _offer_specs(text):
    folded = fold(text or "")
    specs = []
    for match in _SPEC_UNITS.finditer(folded):
        number = match.group(1).replace(",", ".")
        if number.endswith(".0"):
            number = number[:-2]
        label = f"{number} {_SPEC_UNIT_LABEL.get(match.group(2), match.group(2))}"
        if label not in specs:
            specs.append(label)
    for word, label in _SPEC_WORDS:
        if re.search(rf"(?<![a-z0-9]){re.escape(word)}(?![a-z0-9])", folded) and label not in specs:
            specs.append(label)
    return specs[:10]


def _title_tokens(text):
    words = re.findall(r"[a-z0-9]+", fold(text or ""))
    return {word for word in words if len(word) >= 2 and word not in _TITLE_STOP}


HW_ORDER = ("Model", "İşlemci", "Ekran kartı", "RAM", "Depolama", "Ekran", "İşletim sistemi")
HW_NOUN = {"RAM": "RAM", "Depolama": "depolama", "Ekran": "ekran", "Ekran kartı": "ekran kartı"}
_HW_BRANDS = (
    "asus", "acer", "lenovo", "hp", "dell", "msi", "apple", "macbook", "monster", "casper", "huawei", "samsung",
    "lg", "gigabyte", "microsoft", "toshiba", "fujitsu", "excalibur", "philips", "sony", "tcl", "vestel",
    "arcelik", "grundig", "xiaomi", "razer", "alienware", "thinkpad",
)
_CPU_PATTERNS = (
    r"intel\s+core\s+ultra\s+\d\s+\d{3}[a-z]{0,2}",
    r"(?:intel\s+)?(?:core\s+)?\bi[3579][\s-]+\d{4,5}[a-z]{0,2}\b",
    r"(?:amd\s+)?ryzen\s*(?:ai\s+)?\d\s+(?:pro\s+)?\d{3,4}[a-z]{0,3}",
    r"apple\s+m[1-5](?:\s+(?:pro|max|ultra))?",
    r"intel\s+(?:celeron|pentium)\s+\w+",
    r"(?:intel\s+)?xeon\s+[a-z0-9-]+",
    r"snapdragon\s+x\s*(?:elite|plus)?",
)
_GPU_NVIDIA = re.compile(
    r"(?:nvidia\s+)?(?:geforce\s+)?(rtx|gtx)\s?(\d{3,4})(\s?ti)?(\s+\d{1,2}\s?gb)?(\s+gddr\d\w?)?(\s+\d{2,3}\s?w\b)?",
    re.I,
)
_GPU_OTHER = (
    (r"(?:amd\s+)?radeon\s+rx\s?\d{3,4}[a-z]{0,2}", None),
    (r"intel\s+arc(?:\s+a\d{3}m?)?", None),
    (r"(?:intel\s+)?iris\s+xe", "Dahili · Intel Iris Xe"),
    (r"intel\s+uhd(?:\s+graphics)?", "Dahili · Intel UHD"),
    (r"(?:amd\s+)?radeon\s+graphics", "Dahili · AMD Radeon"),
)


_PRETTY_WORDS = {
    "amd": "AMD", "ssd": "SSD", "hdd": "HDD", "nvme": "NVMe", "freedos": "FreeDOS", "intel": "Intel",
    "nvidia": "NVIDIA", "geforce": "GeForce", "oled": "OLED", "qled": "QLED", "ips": "IPS", "fhd": "FHD",
    "uhd": "UHD", "qhd": "QHD", "wuxga": "WUXGA", "tuf": "TUF", "hp": "HP", "msi": "MSI", "lg": "LG",
    "tcl": "TCL", "rog": "ROG", "ddr4": "DDR4", "ddr5": "DDR5", "macbook": "MacBook", "ideapad": "IdeaPad",
    "thinkpad": "ThinkPad", "tasinabilir": "Taşınabilir", "televizyon": "Televizyon", "wifi": "Wi-Fi",
}


def _pretty_words(text):
    words = []
    for word in str(text).split():
        key = word.lower()
        if key in _PRETTY_WORDS:
            words.append(_PRETTY_WORDS[key])
        elif re.fullmatch(r"i[3579]", key):
            words.append(key)
        elif re.fullmatch(r"\d+hz", key):
            words.append(key[:-2] + "Hz")
        elif any(ch.isdigit() for ch in word):
            words.append(word.upper())
        elif word.isupper() and len(word) > 1:
            words.append(word)
        else:
            words.append(word[:1].upper() + word[1:])
    return " ".join(words)


def _hw_gb(number, unit):
    value = float(str(number).replace(",", "."))
    return value * 1024 if unit.lower() == "tb" else value


def _hw_size(number, unit):
    value = float(str(number).replace(",", "."))
    if unit.lower() == "tb":
        return f"{value:g} TB"
    return f"{value:g} GB"


_RAM_GEN = {
    "DDR3": 3, "DDR4": 4, "DDR5": 5, "DDR6": 6,
    "LPDDR3": 3, "LPDDR4": 4, "LPDDR4X": 4.2, "LPDDR5": 5, "LPDDR5X": 5.3,
}


def _ram_generation(kind):
    if not kind:
        return 0
    if kind in _RAM_GEN:
        return _RAM_GEN[kind]
    if kind[-1:].isalpha() and kind[:-1] in _RAM_GEN:
        return _RAM_GEN[kind[:-1]]
    return 0


def _ram_score(gb, kind, speed):
    # Kapasite baskın kalır: 32 GB DDR4, 16 GB DDR5'i geçer.
    # Aynı kapasitede nesil, sonra MHz sırayı belirler.
    return gb + _ram_generation(kind) * 4 + (speed or 0) / 1000


def _ram_phrase(mine, base):
    same = mine.get("gb") and mine.get("gb") == base.get("gb")
    if not same:
        if mine["value"] > base["value"]:
            return f"{mine['text']} RAM ({base['text']} yerine)"
        return f"{mine['text']} RAM (diğerinde {base['text']})"
    bits = []
    mine_gen = _ram_generation(mine.get("kind") or "")
    base_gen = _ram_generation(base.get("kind") or "")
    if mine_gen and base_gen and mine_gen != base_gen:
        bits.append("daha yeni nesil" if mine_gen > base_gen else "daha eski nesil")
    mine_speed = mine.get("speed") or 0
    base_speed = base.get("speed") or 0
    if mine_speed and base_speed and mine_speed != base_speed:
        bits.append("daha yüksek hızda" if mine_speed > base_speed else "daha düşük hızda")
    why = " ve ".join(bits)
    if why:
        return f"{mine['text']} RAM; diğerinde {base['text']}. Kapasite aynı, bellek {why}"
    if mine["value"] > base["value"]:
        return f"{mine['text']} RAM ({base['text']} yerine)"
    return f"{mine['text']} RAM (diğerinde {base['text']})"


def _memory_note(entries):
    rams = [(entry, entry["hw"]["RAM"]) for entry in entries if entry["hw"].get("RAM", {}).get("gb")]
    if len(rams) < 2:
        return ""
    if len({item[1]["gb"] for item in rams}) != 1:
        return ""
    best = max(rams, key=lambda item: item[1]["value"])
    worst = min(rams, key=lambda item: item[1]["value"])
    if best[0] is worst[0] or best[1]["value"] == worst[1]["value"]:
        return ""
    if (best[1].get("kind") or "") == (worst[1].get("kind") or "") and (best[1].get("speed") or 0) == (worst[1].get("speed") or 0):
        return ""
    gb = best[1]["gb"]
    return (
        f"RAM kapasitesi aynı ({gb:g} GB) ama bellek aynı değil: "
        f"{best[0]['label']} {best[1]['text']}, {worst[0]['label']} {worst[1]['text']}. "
        f"Kıyas kapasiteye değil, nesile ve hıza göre yapıldı."
    )


def _spec_hardware_hints(text):
    table = {fold(label): value for label, value in offer_spec_table(text)}

    def pick(*keys):
        for key in keys:
            if table.get(key):
                return table[key]
        return ""

    hints = []
    family = pick("islemci tipi", "islemci", "islemci markasi", "processor", "cpu")
    model = pick("islemci modeli", "islemci model", "processor model")
    if family or model:
        cpu = family if not model or fold(model) in fold(family) else f"{family} {model}"
        hints.append(f"İşlemci: {cpu.strip()}")
    gpu = pick("ekran karti modeli", "ekran karti", "grafik karti", "ekran karti chipseti", "graphics")
    memory = pick("ekran karti hafizasi", "ekran karti bellegi (gb)", "ekran karti bellegi", "ekran karti bellek kapasitesi")
    if gpu:
        hints.append(f"Ekran kartı: {gpu} {memory}".strip())
    return " ".join(hints)


def offer_hardware(text):
    flat = " ".join(f"{text or ''} {_spec_hardware_hints(text)}".split())
    found = {}
    cpu = None
    for pattern in _CPU_PATTERNS:
        cpu = re.search(pattern, flat, re.I)
        if cpu:
            break
    if cpu:
        name = _pretty_words(cpu.group(0))
        name = re.sub(r"(?i)ryzen\s*(\d)", r"Ryzen \1", name)
        name = re.sub(r"(?i)^ryzen", "AMD Ryzen", name)
        core = re.fullmatch(r"(?i)(?:intel\s+)?(?:core\s+)?i([3579])[\s-]+(\d{4,5}[a-z]{0,2})", name)
        if core:
            name = f"Intel Core i{core.group(1)} {core.group(2).upper()}"
        found["İşlemci"] = {"text": name, "short": name, "value": None}
        before = flat[max(0, cpu.start() - 90):cpu.start()]
        brand = None
        for match in re.finditer(rf"(?i)\b({'|'.join(_HW_BRANDS)})\b", before):
            brand = match
        if brand:
            model = before[brand.start():].strip(" -|·,")
            if model == model.lower():
                model = _pretty_words(model)
            if 2 <= len(model.split()) <= 8 and len(model) <= 70:
                found["Model"] = {"text": model, "short": model, "value": None}
    best = None
    for match in _GPU_NVIDIA.finditer(flat):
        detail = sum(1 for index in (3, 4, 5, 6) if match.group(index))
        if best is None or detail > best[0]:
            best = (detail, match)
    if best:
        match = best[1]
        family = match.group(1).upper()
        number = int(match.group(2))
        name = f"{family} {number}" + (" Ti" if match.group(3) else "")
        bits = [f"NVIDIA GeForce {name}"]
        if match.group(4):
            bits.append(match.group(4).strip().upper().replace("GB", " GB").replace("  ", " "))
        if match.group(5):
            bits.append(match.group(5).strip().upper())
        if match.group(6):
            bits.append(re.sub(r"(?i)\s?w$", " W", match.group(6).strip()))
        tier, series = number % 100, number // 100
        score = tier * 100 + series - (5000 if family == "GTX" else 0) + (5 if match.group(3) else 0)
        found["Ekran kartı"] = {"text": " · ".join(bits), "short": name, "value": score}
    else:
        for pattern, label in _GPU_OTHER:
            match = re.search(pattern, flat, re.I)
            if match:
                name = label or _pretty_words(match.group(0))
                found["Ekran kartı"] = {"text": name, "short": name, "value": None}
                break
    pair = re.search(r"(\d{1,3})\s?gb\s+(\d{3,4}\s?gb|\d\s?tb)\s+(?:m\.2\s+|nvme\s+|pcie\s+)*(?:ssd|hdd)", flat, re.I)
    ram = (
        re.search(r"(?i)\bram\b[^:]{0,30}:\s*(\d{1,3})\s?gb", flat)
        or re.search(r"(?i)(?:bellek|memory)[^:]{0,20}:\s*(\d{1,3})\s?gb", flat)
        or re.search(r"(?i)\bram\s*\(sistem belle\w*\)\s*(\d{1,3})\s?gb", flat)
    )
    detail = re.search(r"(\d{1,3})\s?gb\s+(l?pddr\d\w?|ddr\d\w?)(?:[\s-]+(\d{4})(?:\s?mhz)?)?", flat, re.I)
    gb = None
    if ram:
        gb = int(ram.group(1))
    elif detail:
        gb = int(detail.group(1))
    elif pair:
        gb = int(pair.group(1))
    if gb:
        kind = detail.group(2).upper() if detail else ""
        if not kind:
            loose = re.search(r"\b(l?pddr\d\w?|ddr\d\w?)\b", flat, re.I)
            kind = loose.group(1).upper() if loose else ""
        speed = detail.group(3) if detail and detail.group(3) else ""
        if not speed:
            loose = re.search(r"(\d{4})\s?mhz", flat, re.I)
            speed = loose.group(1) if loose else ""
        speed_num = int(speed) if str(speed).isdigit() else 0
        text_ram = " ".join(bit for bit in (f"{gb} GB", kind, f"{speed_num} MHz" if speed_num else "") if bit)
        found["RAM"] = {
            "text": text_ram,
            "short": text_ram,
            "value": _ram_score(gb, kind, speed_num),
            "gb": gb,
            "kind": kind,
            "speed": speed_num,
        }
    labeled = re.search(r"(?i)ssd\s+kapasitesi\s*:?\s*(\d+)\s?(gb|tb)", flat)
    disk = re.search(r"(\d{3,4}\s?gb|\d\s?tb)\s+((?:m\.2\s+|nvme\s+|pcie\s+)*)(ssd|hdd)", flat, re.I)
    size = None
    if labeled:
        size = (labeled.group(1), labeled.group(2), "SSD", "")
    elif disk:
        amount = re.match(r"(\d+)\s?(gb|tb)", disk.group(1), re.I)
        size = (amount.group(1), amount.group(2), disk.group(3).upper(), disk.group(2).strip().upper())
    elif pair:
        amount = re.match(r"(\d+)\s?(gb|tb)", pair.group(2), re.I)
        size = (amount.group(1), amount.group(2), "SSD", "")
    if size:
        shown = _hw_size(size[0], size[1])
        extra = size[3]
        if not extra and disk:
            extra = disk.group(2).strip().upper()
        text_disk = f"{shown} {size[2]}" + (f" · {extra}" if extra else "")
        found["Depolama"] = {"text": text_disk, "short": shown, "value": _hw_gb(size[0], size[1])}
    screen = []
    inch = re.search(r"(?<![\d.,])(\d{2}(?:[.,]\d)?)\s?(?:\"|”|″|inç|inc|inch)(?![a-z])", flat, re.I)
    if inch:
        screen.append(f"{inch.group(1).replace(',', '.')} inç")
    folded = fold(flat)
    for word, label in (("wuxga", "WUXGA"), ("wqxga", "WQXGA"), ("qhd", "QHD"), ("fhd", "FHD"), ("full hd", "Full HD"), ("2k", "2K"), ("4k", "4K"), ("uhd", "UHD")):
        if re.search(rf"(?<![a-z0-9]){re.escape(word)}(?![a-z0-9])", folded):
            screen.append(label)
            break
    for word, label in (("oled", "OLED"), ("mini led", "Mini LED"), ("qled", "QLED"), ("ips", "IPS")):
        if re.search(rf"(?<![a-z0-9]){re.escape(word)}(?![a-z0-9])", folded):
            screen.append(label)
            break
    hz = re.search(r"(?<!\d)(\d{2,3})\s?hz\b", flat, re.I)
    if hz:
        screen.append(f"{hz.group(1)} Hz")
    if screen:
        found["Ekran"] = {"text": " · ".join(screen), "short": f"{hz.group(1)} Hz" if hz else screen[0], "value": float(hz.group(1)) if hz else None}
    system = None
    if re.search(r"(?i)\bfree\s?dos\b", flat):
        system = "FreeDOS (Windows yok)"
    else:
        win = re.search(r"(?i)windows\s*(11|10)(\s+(?:home|pro))?", flat)
        if win:
            system = f"Windows {win.group(1)}" + (f" {win.group(2).strip().title()}" if win.group(2) else "")
        elif re.search(r"(?i)\bmacos\b", flat):
            system = "macOS"
        elif re.search(r"(?i)\b(ubuntu|linux)\b", flat):
            system = "Linux"
    if system:
        found["İşletim sistemi"] = {"text": system, "short": system, "value": None}
    return found


def offer_product_match(entries):
    models = [entry["hw"].get("Model", {}).get("text") for entry in entries]
    if len(entries) >= 2 and all(models):
        if len({fold(model) for model in models}) == 1:
            return {"level": "same", "short": "Aynı model", "text": f"Tüm teklifler aynı modeli veriyor: {models[0]}. Fark yalnız fiyat ve koşullarda."}
        listed = "; ".join(f"{entry['label']}: {model}" for entry, model in zip(entries, models))
        return {
            "level": "similar",
            "short": "Farklı model",
            "text": f"Aynı ihtiyaç, farklı modeller. {listed}. Fiyat farkının bir kısmı donanım farkından geliyor; işlemci, ekran kartı, RAM ve depolama aşağıdaki tabloda parça parça kıyaslandı.",
        }
    titles = [f"{entry['row']['title'] or ''} {entry['row']['notes'] or ''}".strip() for entry in entries]
    tokens = [_title_tokens(title) for title in titles]
    if len(entries) < 2 or any(not item for item in tokens):
        return {
            "level": "unknown",
            "short": "Belirsiz",
            "text": "Teklif adlarından ürün tanımı çıkmıyor. Kıyas yalnız fiyat ve koşullar üzerinden yapıldı; aynı ürün olduğu tedarikçiden teyit edilmeli.",
        }
    common = set.intersection(*tokens)
    union = set.union(*tokens)
    overlap = len(common) / len(union) if union else 0
    title_specs = [set(_offer_specs(title)) for title in titles]
    spec_common = set.intersection(*title_specs) if title_specs else set()
    spec_diff = set.union(*title_specs) - spec_common if title_specs else set()
    first = titles[0].split()
    shown = " ".join(word for word in first if (_title_tokens(word) & common)) or ", ".join(sorted(common))
    if overlap >= 0.6 and not spec_diff:
        text = f"Aynı ürün tanımı: {shown}."
        if len(spec_common) < 2:
            text += " Teklif adında model ve teknik föy yok; marka ya da model farkı fiyata yansıyor olabilir. Karar öncesi model kodu istenmeli."
        return {"level": "same", "short": "Aynı ürün", "text": text}
    if overlap >= 0.3 or spec_common:
        text = f"Aynı ihtiyaç, farklı model ya da kapsam. Ortak: {shown}."
        if spec_diff:
            text += f" Ayrışan özellik: {', '.join(sorted(spec_diff))}."
        text += " Fiyat farkının bir kısmı ürün farkından gelir."
        return {"level": "similar", "short": "Benzer ürün", "text": text}
    return {
        "level": "different",
        "short": "Farklı ürün",
        "text": "Teklifler farklı ürünler tarif ediyor. Fiyatlar birebir kıyaslanmamalı; önce ihtiyaç tek bir şartnameye bağlanmalı.",
    }


def _ratio_low(values, value):
    present = [item for item in values if item is not None and item > 0]
    if value is None or value <= 0 or not present:
        return 0.0
    return min(present) / value


def _ratio_high(values, value):
    present = [item for item in values if item is not None]
    top = max(present) if present else 0
    if value is None or top <= 0:
        return 0.0
    return value / top


def _months_text(months):
    if not months:
        return ""
    return f"{months // 12} yıl" if months % 12 == 0 else f"{months} ay"


_SPEC_SKIP_ROWS = re.compile(
    r"\b(?:renk|model kodu|urun modeli|kullanim amaci|garanti|mensei|urun guvenligi|klavye|ram|sistem bellegi|"
    r"ssd|hdd|harddisk|hard disk|disk turu|depolama|islemci tipi|islemci modeli|ekran karti|ekran boyutu|"
    r"isletim sistemi|yenileme)\b|^islemci$"
)
_SPEC_ALIASES = {
    "ekran cozunurlugu": "cozunurluk", "max ekran cozunurlugu": "cozunurluk", "cozunurluk": "cozunurluk",
    "sarjli kullanim suresi": "pil omru", "batarya omru": "pil omru", "pil omru": "pil omru",
    "cihaz agirligi": "agirlik", "agirlik": "agirlik", "urun agirligi": "agirlik",
    "panel tipi": "panel tipi", "ekran panel tipi": "panel tipi",
    "islemci frekansi": "islemci hizi", "maksimum islemci hizi": "islemci hizi", "islemci hizi": "islemci hizi",
    "islemci cekirdek sayisi": "cekirdek sayisi", "cekirdek sayisi": "cekirdek sayisi",
    "bluetooth": "bluetooth", "bluetooth ozelligi": "bluetooth",
    "kamera": "kamera", "webcam": "kamera",
}


def spec_key(label):
    key = " ".join(re.sub(r"\([^)]*\)", " ", fold(label)).split())
    return _SPEC_ALIASES.get(key, key)


def _spec_lookup(pairs):
    table = {}
    for label, value in pairs:
        table.setdefault(spec_key(label), (re.sub(r"\s*\([^)]*\)", "", label).strip() or label, value))
    return table


def shared_spec_rows(entries, limit=14):
    if sum(1 for entry in entries if entry["table"]) < 2:
        return []
    counts, order = {}, []
    for entry in entries:
        for key, (label, _value) in entry["table"].items():
            if key not in counts:
                order.append((key, label))
            counts[key] = counts.get(key, 0) + 1
    rows = []
    for key, label in order:
        if counts[key] < 2 or _SPEC_SKIP_ROWS.search(key):
            continue
        values = {fold(entry["table"][key][1]) for entry in entries if key in entry["table"]}
        if len(values) < 2 and counts[key] == len(entries):
            continue
        rows.append((key, label))
        if len(rows) >= limit:
            break
    return rows


def compare_offers(project, rows, mode):
    weights = COMPARE_MODES.get(mode, COMPARE_MODES["Dengeli"])
    vendor_count = {}
    for row in rows:
        key = fold(row["vendor"] or "")
        vendor_count[key] = vendor_count.get(key, 0) + 1
    entries = []
    for row in rows:
        amount = float(row["amount"]) if row["amount"] is not None else None
        text = (row["source_text"] if "source_text" in row.keys() else "") or ""
        url = (row["url"] if "url" in row.keys() else "") or ""
        if url and "Adres: " not in text:
            text = "\n".join(part for part in (text, f"Adres: {offer_url_slug_text(url)}") if part)
        profile = extract_offer_profile(row["vendor"], amount, text, bool(row["file_path"] or url))
        effective = profile["amount"]
        if effective is not None and profile["kdv"] == "haric":
            effective = round(effective * 1.2, 2)
        label = row["vendor"] or "Firma"
        if vendor_count.get(fold(row["vendor"] or ""), 0) > 1 and row["title"]:
            label = f"{label} · {row['title']}"
        delivery = profile["delivery_days"]
        if delivery is None and profile["stock"]:
            delivery = 1
        entries.append({
            "row": row,
            "label": label,
            "profile": profile,
            "amount": profile["amount"],
            "effective": effective,
            "delivery": delivery,
            "scope": len(profile["includes"]) - len(profile["excludes"]),
            "specs": _offer_specs(f"{row['title'] or ''} {row['notes'] or ''} {text}"),
            "hw": offer_hardware(f"{row['title'] or ''}\n{row['notes'] or ''}\n{text}"),
            "table": _spec_lookup(offer_spec_table(text)),
        })
    prices = [entry["effective"] for entry in entries]
    warranties = [entry["profile"]["warranty_months"] for entry in entries]
    deliveries = [entry["delivery"] for entry in entries]
    payments = [entry["profile"]["payment_days"] for entry in entries]
    scopes = [entry["scope"] for entry in entries]
    has_scope = any(entry["profile"]["includes"] or entry["profile"]["excludes"] for entry in entries)
    hw_keys = []
    for key in ("Ekran kartı", "RAM", "Depolama", "Ekran"):
        values = [entry["hw"].get(key, {}).get("value") for entry in entries]
        present = [value for value in values if value is not None]
        if len(present) >= 2 and len(set(present)) > 1:
            hw_keys.append(key)
    for entry in entries:
        entry["hw_score"] = _hardware_score(entry, entries, hw_keys)
    active = {
        "hardware": bool(hw_keys),
        "price": any(value is not None for value in prices),
        "warranty": any(warranties),
        "delivery": any(deliveries),
        "payment": any(value for value in payments if value is not None),
        "scope": has_scope and len(set(scopes)) > 1,
    }
    used = {key: weight for key, weight in weights.items() if active[key]}
    total = sum(used.values()) or 1.0
    used = {key: weight / total for key, weight in used.items()}
    low_scope, high_scope = (min(scopes), max(scopes)) if scopes else (0, 0)
    for entry in entries:
        parts = {
            "hardware": entry["hw_score"],
            "price": _ratio_low(prices, entry["effective"]),
            "warranty": _ratio_high(warranties, entry["profile"]["warranty_months"]),
            "delivery": _ratio_low(deliveries, entry["delivery"]),
            "payment": _ratio_high(payments, entry["profile"]["payment_days"]),
            "scope": (entry["scope"] - low_scope) / (high_scope - low_scope) if high_scope > low_scope else 1.0,
        }
        entry["parts"] = {key: parts[key] for key in used}
        entry["score"] = round(100 * sum(parts[key] * weight for key, weight in used.items())) if used else 0
    ranked = sorted(entries, key=lambda entry: (-entry["score"], entry["effective"] if entry["effective"] is not None else float("inf")))
    priced = [entry for entry in entries if entry["effective"] is not None]
    cheapest = min(priced, key=lambda entry: entry["effective"]) if priced else None
    dearest = max(priced, key=lambda entry: entry["effective"]) if priced else None
    budget = float(project["budget"]) if project["budget"] is not None else None
    for entry in entries:
        entry["value"] = _value_line(entry, cheapest)
    match = offer_product_match(entries)
    return {
        "entries": entries,
        "ranked": ranked,
        "cheapest": cheapest,
        "dearest": dearest,
        "budget": budget,
        "weights": used,
        "active": active,
        "match": match,
        "verdict": _compare_verdict(ranked, cheapest, budget, used, match),
        "gaps": _compare_gaps(entries, budget, match),
    }


def _hardware_score(entry, entries, keys):
    if not keys:
        return 0.0
    total = 0.0
    for key in keys:
        values = [item["hw"].get(key, {}).get("value") for item in entries]
        mine = entry["hw"].get(key, {}).get("value")
        present = [value for value in values if value is not None]
        if mine is None:
            continue
        if key == "Ekran kartı":
            low, high = min(present), max(present)
            total += (mine - low) / (high - low) if high > low else 1.0
        else:
            total += mine / max(present) if max(present) else 0.0
    return total / len(keys)


def _gains_over(entry, base):
    gains = []
    losses = []
    for key in ("Ekran kartı", "RAM", "Depolama", "Ekran"):
        mine_hw, base_hw = entry["hw"].get(key), base["hw"].get(key)
        if not mine_hw or not base_hw or mine_hw.get("value") is None or base_hw.get("value") is None:
            continue
        if mine_hw["value"] == base_hw["value"]:
            continue
        if key == "RAM":
            phrase = _ram_phrase(mine_hw, base_hw)
        elif mine_hw["value"] > base_hw["value"]:
            phrase = f"{mine_hw['short']} {HW_NOUN[key]} ({base_hw['short']} yerine)"
        else:
            phrase = f"{mine_hw['short']} {HW_NOUN[key]} (diğerinde {base_hw['short']})"
        if mine_hw["value"] > base_hw["value"]:
            gains.append(phrase)
        else:
            losses.append(phrase)
    mine, theirs = entry["profile"], base["profile"]
    if (mine["warranty_months"] or 0) > (theirs["warranty_months"] or 0):
        gains.append(f"{_months_text((mine['warranty_months'] or 0) - (theirs['warranty_months'] or 0))} daha uzun garanti")
    elif (theirs["warranty_months"] or 0) > (mine["warranty_months"] or 0):
        losses.append(f"{_months_text((theirs['warranty_months'] or 0) - (mine['warranty_months'] or 0))} daha kısa garanti")
    if entry["delivery"] and base["delivery"]:
        if entry["delivery"] < base["delivery"]:
            gains.append(f"{base['delivery'] - entry['delivery']} gün daha erken teslim")
        elif entry["delivery"] > base["delivery"]:
            losses.append(f"{entry['delivery'] - base['delivery']} gün daha geç teslim")
    elif entry["delivery"] and not base["delivery"]:
        gains.append("teslim süresi yazılı")
    if (mine["payment_days"] or 0) > (theirs["payment_days"] or 0):
        gains.append(f"{(mine['payment_days'] or 0) - (theirs['payment_days'] or 0)} gün daha uzun vade")
    elif (theirs["payment_days"] or 0) > (mine["payment_days"] or 0):
        losses.append(f"{(theirs['payment_days'] or 0) - (mine['payment_days'] or 0)} gün daha kısa vade")
    extra = [label for word, label in OFFER_FEATURES if word in mine["includes"] and word not in theirs["includes"]]
    if extra:
        gains.append(", ".join(extra) + " dahil")
    missing = [label for word, label in OFFER_FEATURES if word in theirs["includes"] and word not in mine["includes"]]
    if missing:
        losses.append(", ".join(missing) + " dahil değil")
    if mine["stock"] and not theirs["stock"]:
        gains.append("stokta")
    return gains, losses


def _value_line(entry, cheapest):
    if cheapest is None or entry["effective"] is None:
        return "Tutar yok; fiyat/fayda hesaplanamadı."
    if entry is cheapest:
        return "En düşük maliyetli teklif. Diğerleri bu tutarın üstüne ne verdiğiyle ölçülür."
    diff = entry["effective"] - cheapest["effective"]
    pct = diff / cheapest["effective"] * 100 if cheapest["effective"] else 0
    gains, losses = _gains_over(entry, cheapest)
    head = f"{cheapest['label']} teklifinden {format_amount(diff)} ₺ (%{pct:.0f}) pahalı."
    if not gains:
        tail = " Bu fark karşılığında belgede ek bir fayda görünmüyor."
    else:
        tail = " Karşılığında: " + "; ".join(gains) + "."
        extra_months = (entry["profile"]["warranty_months"] or 0) - (cheapest["profile"]["warranty_months"] or 0)
        if extra_months > 0:
            tail += f" Ek garantinin aylık maliyeti yaklaşık {format_amount(diff / extra_months)} ₺."
    if losses:
        tail += " Buna rağmen: " + "; ".join(losses) + "."
    return head + tail


def _compare_verdict(ranked, cheapest, budget, used, match):
    if not ranked:
        return {"pick": None, "text": "Karşılaştırılacak teklif yok."}
    top = ranked[0]
    reasons = []
    if top is cheapest:
        reasons.append("en düşük maliyet")
    elif cheapest is not None and top["effective"] is not None:
        gains, _losses = _gains_over(top, cheapest)
        diff = top["effective"] - cheapest["effective"]
        reasons.append(f"{format_amount(diff)} ₺ daha pahalı olmasına rağmen " + ("; ".join(gains) if gains else "koşulları daha iyi"))
    best_parts = [
        COMPARE_LABELS[key] for key in used
        if key != "price" and top["parts"].get(key, 0) >= 0.999 and any(entry["parts"].get(key, 0) < 0.999 for entry in ranked)
    ]
    if best_parts:
        reasons.append("en iyi " + ", ".join(best_parts).lower())
    text = f"{top['label']} öne çıkıyor: " + "; ".join(reasons) + "."
    memory = _memory_note(ranked)
    if memory:
        text += " " + memory
    if set(used) == {"price"}:
        text += " Belgelerde garanti, teslim, vade gibi koşullar okunmadı; öneri yalnız fiyata dayanıyor."
    if len(ranked) > 1 and ranked[0]["score"] - ranked[1]["score"] < 5:
        text += f" {ranked[1]['label']} ile fark {ranked[0]['score'] - ranked[1]['score']} puan; karar koşul pazarlığıyla netleşir."
    if budget is not None and top["effective"] is not None and top["effective"] > budget:
        inside = [entry for entry in ranked if entry["effective"] is not None and entry["effective"] <= budget]
        text += f" Dikkat: bütçeyi {format_amount(top['effective'] - budget)} ₺ aşıyor."
        if inside:
            text += f" Bütçe içindeki en iyi seçenek {inside[0]['label']}."
    if match["level"] == "different":
        text += " Ürünler farklı olduğu için bu öneri ön değerlendirmedir."
    return {"pick": top, "text": text}


def _compare_gaps(entries, budget, match):
    gaps = []
    checks = (
        ("Garanti süresi", lambda entry: entry["profile"]["warranty_months"]),
        ("Teslim süresi", lambda entry: entry["delivery"]),
        ("Ödeme vadesi", lambda entry: entry["profile"]["payment_days"] is not None),
        ("KDV durumu", lambda entry: entry["profile"]["kdv"]),
        ("Teklif geçerlilik süresi", lambda entry: entry["profile"]["validity_days"]),
    )
    for label, has in checks:
        missing = [entry["label"] for entry in entries if not has(entry)]
        if not missing:
            continue
        if len(missing) == len(entries):
            gaps.append(f"{label} hiçbir teklifte yazmıyor. Tüm tedarikçilerden yazılı istenmeli.")
        else:
            gaps.append(f"{label} {', '.join(missing)} teklifinde yok; eşit kıyas için istenmeli.")
    for entry in entries:
        if entry["profile"]["kdv"] == "haric" and entry["amount"] is not None:
            gaps.append(f"{entry['label']} KDV hariç vermiş. Kıyasta %20 eklenerek {format_amount(entry['effective'])} ₺ alındı.")
        has_url = "url" in entry["row"].keys() and entry["row"]["url"]
        if not entry["row"]["file_path"] and not has_url:
            gaps.append(f"{entry['label']} için teklif belgesi ya da sayfası eklenmemiş; tutar elle girilmiş.")
        elif not entry["row"]["file_path"]:
            gaps.append(f"{entry['label']} yalnız internet sayfasıyla eklendi. Garanti ve teslim gibi koşullar için yazılı teklif istenmeli.")
        elif not entry["profile"]["has_text"]:
            gaps.append(f"{entry['label']} belgesinden metin okunamadı; koşullar kıyasa girmedi.")
    if budget is not None:
        over = [entry["label"] for entry in entries if entry["effective"] is not None and entry["effective"] > budget]
        if over:
            gaps.append(f"Bütçeyi aşan: {', '.join(over)}.")
    if match["level"] in {"same", "unknown"} and len(entries) > 1:
        specs = [set(entry["specs"]) for entry in entries]
        if not any(specs):
            gaps.append("Hiçbir teklifte model kodu veya teknik özellik okunmadı. Aynı ürünün fiyatı mı kıyaslanıyor, teyit edilmeli.")
    return gaps[:10]


def compose_offer_text(typed, lines):
    notes = []
    for line in str(typed or "").splitlines():
        stripped = line.strip()
        if stripped and not stripped.startswith(("•", "-")):
            notes.append(stripped)
    body = [f"• {line}" for line in lines if line]
    text = "\n".join(notes + body)
    return text[:2000]


def preview_offer_text(vendor, amount, text, siblings):
    profile = extract_offer_profile(vendor, amount, text, True)
    profiles = [profile]
    for row in siblings:
        cached = row["source_text"] if "source_text" in row.keys() else ""
        other_amount = float(row["amount"]) if row["amount"] is not None else None
        profiles.append(extract_offer_profile(row["vendor"], other_amount, cached or "", bool(row["file_path"])))
    pros, cons = narrate_offer(profile, profiles)
    return "\n".join(f"• {line}" for line in pros), "\n".join(f"• {line}" for line in cons)


def store_report_file(src):
    source = Path(src)
    if not source.is_file() or source.is_symlink():
        raise ValueError("Dosya kullanılamıyor.")
    resolved = source.resolve()
    if resolved.suffix.lower() not in REPORT_EXT or not report_signature_ok(resolved):
        raise ValueError("Yalnızca ekran görüntüsü, .eml e-posta veya PDF eklenebilir.")
    try:
        if resolved.stat().st_size <= 0 or resolved.stat().st_size > MAX_BYTES:
            raise ValueError("Dosya boş ya da 20 MB sınırının üstünde.")
    except OSError as exc:
        raise ValueError("Dosya okunamadı.") from exc
    if is_within(resolved, REPORT_DIR):
        return str(resolved)
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    os.chmod(REPORT_DIR, 0o700)
    dest = REPORT_DIR / f"{datetime.now().strftime('%Y%m%d-%H%M%S')}-{secrets.token_hex(3)}-{safe_name(resolved.name)}"
    shutil.copy2(resolved, dest)
    os.chmod(dest, 0o600)
    return str(dest)


_REPORT_LOCK = threading.Lock()


def report_tool():
    with _REPORT_LOCK:
        return _report_tool()


def _report_tool():
    if sys.platform != "darwin":
        raise ValueError("PDF birleştirme aracı yalnız macOS üzerinde hazır.")
    source = Path(__file__).resolve().parent / "İş Merkezi.app" / "Contents" / "Resources" / "rapor-pdf.swift"
    if not source.is_file():
        raise ValueError("PDF birleştirme aracı eksik.")
    tool_dir = DATA_DIR / "araclar"
    tool = tool_dir / "rapor-pdf"
    stamp = str(source.stat().st_mtime_ns)
    marker = tool_dir / "rapor-pdf.stamp"
    if tool.is_file() and marker.is_file() and marker.read_text(encoding="utf-8") == stamp:
        return tool
    tool_dir.mkdir(parents=True, exist_ok=True)
    os.chmod(tool_dir, 0o700)
    result = subprocess.run(["swiftc", "-O", "-o", str(tool), str(source)], capture_output=True, text=True)
    if result.returncode != 0 or not tool.is_file():
        raise ValueError(result.stderr.strip() or "PDF aracı derlenemedi.")
    marker.write_text(stamp, encoding="utf-8")
    os.chmod(tool, 0o700)
    return tool


def note_tool():
    if sys.platform != "darwin":
        raise ValueError("Not PDF’i yalnız Mac üzerinde hazırlanır.")
    with _REPORT_LOCK:
        source = Path(__file__).resolve().parent / "İş Merkezi.app" / "Contents" / "Resources" / "not-pdf.swift"
        if not source.is_file():
            raise ValueError("Not PDF aracı eksik.")
        tool_dir = DATA_DIR / "araclar"
        tool = tool_dir / "not-pdf"
        stamp = str(source.stat().st_mtime_ns)
        marker = tool_dir / "not-pdf.stamp"
        if tool.is_file() and marker.is_file() and marker.read_text(encoding="utf-8") == stamp:
            return tool
        tool_dir.mkdir(parents=True, exist_ok=True)
        os.chmod(tool_dir, 0o700)
        result = subprocess.run(["swiftc", "-O", "-o", str(tool), str(source)], capture_output=True, text=True)
        if result.returncode != 0 or not tool.is_file():
            raise ValueError(result.stderr.strip() or "Not PDF aracı derlenemedi.")
        marker.write_text(stamp, encoding="utf-8")
        os.chmod(tool, 0o700)
        return tool


def write_note_pdf(path, title, blocks):
    path = Path(path).expanduser()
    if path.suffix.lower() != ".pdf":
        path = path.with_suffix(".pdf")
    if path.is_symlink() or not path.parent.is_dir() or path.parent.is_symlink():
        raise ValueError("PDF kayıt yeri geçersiz.")
    job = {
        "output": str(path),
        "title": one_line(title, 120) or "Görev notu",
        "root": str(NOTE_DIR.resolve()),
        "width": 820,
        "height": 800,
        "items": [],
        "flow": True,
        "blocks": blocks,
    }
    tool_dir = DATA_DIR / "araclar"
    tool_dir.mkdir(parents=True, exist_ok=True)
    job_path = tool_dir / "not-job.json"
    job_path.write_text(json.dumps(job, ensure_ascii=False), encoding="utf-8")
    result = subprocess.run([str(note_tool()), str(job_path)], capture_output=True, text=True, timeout=40)
    if result.returncode != 0 or not path.is_file():
        raise ValueError("Not PDF olarak yazılamadı.")
    return path


NOTE_COLORS = ("#17202a", "#0e8f6c", "#8d2f2f", "#1d4e89", "#8a5a00")
NOTE_HILITES = ("#FFF3A3", "#C8F5D4", "#F8C9C9", "#D5E6FF")
NOTE_SIZES = (14, 18, 24, 32)
NOTE_IMAGE_EXT = {".png", ".jpg", ".jpeg", ".gif", ".webp", ".tif", ".tiff", ".heic"}


def load_note_items(raw):
    try:
        data = json.loads(raw or "")
    except (json.JSONDecodeError, TypeError):
        return []
    items = data.get("items") if isinstance(data, dict) else None
    if not isinstance(items, list):
        return []
    clean = []
    for item in items[:60]:
        if not isinstance(item, dict):
            continue
        kind = item.get("type")
        color = item.get("color") if item.get("color") in NOTE_COLORS else NOTE_COLORS[0]
        try:
            x = max(0, min(float(item.get("x", 24)), 2000))
            y = max(0, min(float(item.get("y", 24)), 4000))
        except (TypeError, ValueError):
            continue
        if kind == "text":
            text = one_line(item.get("text"), 400)
            if not text:
                continue
            size = int(item.get("size") or 18)
            if size not in NOTE_SIZES:
                size = 18
            clean.append({"type": "text", "x": x, "y": y, "text": text, "size": size, "color": color})
        elif kind == "shape" and item.get("kind") in {"rect", "oval", "line"}:
            try:
                x2 = max(0, min(float(item.get("x2", x)), 2000))
                y2 = max(0, min(float(item.get("y2", y)), 4000))
            except (TypeError, ValueError):
                continue
            clean.append({"type": "shape", "kind": item["kind"], "x": x, "y": y, "x2": x2, "y2": y2, "color": color})
        elif kind == "image":
            path = Path(str(item.get("path") or ""))
            if not path.is_file() or path.is_symlink() or not is_within(path, NOTE_DIR):
                continue
            if not file_signature_ok(path) and path.suffix.lower() != ".gif":
                continue
            if path.suffix.lower() == ".gif" and not path.read_bytes()[:6] in {b"GIF87a", b"GIF89a"}:
                continue
            try:
                w = max(8, min(float(item.get("w", 40)), 800))
                h = max(8, min(float(item.get("h", 40)), 800))
            except (TypeError, ValueError):
                continue
            clean.append({"type": "image", "x": x, "y": y, "path": str(path.resolve()), "w": w, "h": h})
    return clean


def _note_data(raw):
    try:
        data = json.loads(raw or "")
    except (json.JSONDecodeError, TypeError):
        return {}
    return data if isinstance(data, dict) else {}


def note_plain(raw):
    data = _note_data(raw)
    body = data.get("body")
    if isinstance(body, str) and body.strip():
        return body
    parts = [str(block.get("text") or "") for block in data.get("blocks") or [] if isinstance(block, dict) and block.get("text")]
    if parts:
        return "\n".join(parts)
    return " ".join(item["text"] for item in load_note_items(raw) if item.get("type") == "text")


def note_has_items(raw):
    if note_plain(raw).strip():
        return True
    data = _note_data(raw)
    for block in data.get("blocks") or []:
        if isinstance(block, dict) and block.get("type") in {"image", "shape"}:
            return True
    return any(item.get("type") in {"image", "shape"} for item in load_note_items(raw))


def remove_note_dir(task_id):
    folder = NOTE_DIR / f"gorev-{int(task_id)}"
    if folder.is_dir() and not folder.is_symlink() and is_within(folder, NOTE_DIR):
        shutil.rmtree(folder)


def store_note_image(task_id, src):
    src = Path(src).expanduser()
    if not src.is_file() or src.is_symlink():
        raise ValueError("Görsel açılamadı.")
    ext = src.suffix.lower()
    if ext not in NOTE_IMAGE_EXT:
        raise ValueError("PNG, JPG, GIF, WEBP veya HEIC seç.")
    if src.stat().st_size > MAX_BYTES or src.stat().st_size == 0:
        raise ValueError("Görsel 20 MB sınırını aşıyor.")
    folder = NOTE_DIR / f"gorev-{int(task_id)}"
    folder.mkdir(parents=True, exist_ok=True)
    os.chmod(NOTE_DIR, 0o700)
    os.chmod(folder, 0o700)
    dest = folder / f"{secrets.token_hex(6)}.png"
    if ext == ".gif":
        dest = dest.with_suffix(".gif")
        shutil.copyfile(src, dest)
    elif ext == ".png":
        shutil.copyfile(src, dest)
    elif sys.platform == "darwin":
        result = subprocess.run(["sips", "-s", "format", "png", str(src), "--out", str(dest)], capture_output=True, text=True, timeout=25)
        if result.returncode != 0 or not dest.is_file():
            dest.unlink(missing_ok=True)
            raise ValueError("Görsel dönüştürülemedi.")
    else:
        raise ValueError("Bu görsel türü yalnız Mac’te açılır.")
    os.chmod(dest, 0o600)
    if dest.suffix.lower() == ".gif":
        ok = dest.read_bytes()[:6] in {b"GIF87a", b"GIF89a"}
    else:
        ok = file_signature_ok(dest)
    if not ok or dest.stat().st_size > MAX_BYTES or not is_within(dest, folder):
        dest.unlink(missing_ok=True)
        raise ValueError("Görsel kabul edilmedi.")
    return dest


def write_report_pdf(path, row, reason, files):
    path = Path(path).expanduser()
    if path.suffix.lower() != ".pdf":
        path = path.with_suffix(".pdf")
    if path.is_symlink() or not path.parent.is_dir() or path.parent.is_symlink():
        raise ValueError("PDF kayıt yeri geçersiz.")
    chosen = []
    for item in files:
        source = Path(item)
        if source.is_file() and not source.is_symlink() and report_signature_ok(source):
            chosen.append(str(source.resolve()))
    amount = format_amount(row["amount"])
    job = {
        "output": str(path),
        "party": row["party"] or "Fatura",
        "meta": f"{amount + ' ₺' if amount else 'tutar yok'} · {row['invoice_no'] or 'no yok'} · {row['invoice_date'] or ''}",
        "reason": reason,
        "files": chosen,
    }
    job_path = DATA_DIR / "araclar" / "rapor-job.json"
    job_path.parent.mkdir(parents=True, exist_ok=True)
    job_path.write_text(json.dumps(job, ensure_ascii=False), encoding="utf-8")
    result = subprocess.run([str(report_tool()), str(job_path)], capture_output=True, text=True, timeout=90)
    if result.returncode != 0 or not path.is_file():
        detail = (result.stderr or result.stdout or "PDF birleştirilemedi.").strip()
        raise ValueError(detail)
    return path


def list_invoice_files(root):
    root = Path(root).resolve()
    found = []
    for dirpath, dirnames, filenames in os.walk(root, followlinks=False):
        dirnames[:] = [
            name
            for name in dirnames
            if not name.startswith(".")
            and name != "__pycache__"
            and not (Path(dirpath) / name).is_symlink()
        ]
        for name in filenames:
            if name.startswith(".") or len(found) >= MAX_SCAN_FILES:
                continue
            path = Path(dirpath) / name
            if path.is_symlink() or path.suffix.lower() not in ALLOWED_EXT:
                continue
            try:
                resolved = path.resolve()
                stat = path.stat()
            except OSError:
                continue
            if stat.st_size <= 0 or stat.st_size > MAX_BYTES:
                continue
            if not is_within(resolved, root) or len(resolved.relative_to(root).parts) > 6:
                continue
            if not file_signature_ok(resolved):
                continue
            found.append(resolved)
    return found


def guess_party(path, root):
    if path.parent.resolve() == root.resolve():
        return path.stem[:120]
    parts = path.parent.name.split("-")
    if len(parts) >= 3 and parts[-1] and not parts[-1].isdigit():
        return parts[-1][:120]
    return path.parent.name[:120]


def path_key(text):
    candidate = Path(str(text or ""))
    try:
        if candidate.exists():
            return str(candidate.resolve())
    except OSError:
        pass
    return str(candidate)


def _sheet_target(files):
    workbook = files["xl/workbook.xml"].decode("utf-8", "replace")
    rels = files["xl/_rels/workbook.xml.rels"].decode("utf-8", "replace")
    match = re.search(r'<sheet[^>]*name="Faturalar"[^>]*r:id="([^"]+)"', workbook)
    if not match:
        match = re.search(r'<sheet[^>]*r:id="([^"]+)"[^>]*name="Faturalar"', workbook)
    if not match:
        raise ValueError("Excel dosyasında Faturalar sayfası yok.")
    rid = match.group(1)
    rel = re.search(rf'<Relationship[^>]*Id="{re.escape(rid)}"[^>]*Target="([^"]+)"', rels)
    if not rel:
        rel = re.search(rf'<Relationship[^>]*Target="([^"]+)"[^>]*Id="{re.escape(rid)}"', rels)
    if not rel:
        raise ValueError("Excel sayfa bağlantısı okunamadı.")
    target = rel.group(1)
    if target.startswith("/"):
        return target.lstrip("/")
    if target.startswith("xl/"):
        return target
    return str(Path("xl") / target)


def _table_target(files, sheet_path):
    rels_name = str(Path(sheet_path).parent / "_rels" / f"{Path(sheet_path).name}.rels")
    if rels_name not in files:
        return None
    text = files[rels_name].decode("utf-8", "replace")
    match = re.search(r'Target="([^"]*table[^"]+\.xml)"', text)
    if not match:
        return None
    target = match.group(1).lstrip("/")
    if target.startswith("xl/"):
        return posixpath.normpath(target)
    return posixpath.normpath(posixpath.join(posixpath.dirname(sheet_path), target))


def _row_cells(body):
    cells = {}
    pattern = re.compile(r'<c r="([A-Z]+)\d+"([^>]*?)(?:/>|>(.*?)</c>)', re.S)
    for match in pattern.finditer(body):
        inner = match.group(3) or ""
        value = ""
        number = re.search(r"<v>(.*?)</v>", inner, re.S)
        text = re.search(r"<t[^>]*>(.*?)</t>", inner, re.S)
        if number:
            value = html.unescape(number.group(1))
        elif text:
            value = html.unescape(text.group(1))
        cells[match.group(1)] = value
    return cells


def read_excel_invoices(path):
    path = Path(path)
    if not path.is_file():
        return []
    with zipfile.ZipFile(path) as workbook:
        files = {name: workbook.read(name) for name in workbook.namelist()}
    sheet = files[_sheet_target(files)].decode("utf-8", "replace")
    rows = []
    for match in re.finditer(r'<row r="(\d+)"[^>]*>(.*?)</row>', sheet, re.S):
        if int(match.group(1)) == 1:
            continue
        cells = _row_cells(match.group(2))
        if not cells.get("B") and not cells.get("A"):
            continue
        total = cells.get("K") or ""
        rows.append(
            {
                "folder": cells.get("A", ""),
                "path": cells.get("B", ""),
                "filename": cells.get("C", ""),
                "doc_no": cells.get("D", ""),
                "invoice_date": from_excel_serial(cells.get("E", "")) if cells.get("E") else "",
                "seller": cells.get("F", ""),
                "tax_no": cells.get("G", ""),
                "base": cells.get("H") or None,
                "vat": cells.get("I") or None,
                "currency": cells.get("J") or "TRY",
                "total": float(total) if total not in ("", None) else None,
                "excel_status": cells.get("L", ""),
            }
        )
    return rows


def _column_attrs(sheet_xml):
    rows = re.findall(r'<row r="(\d+)"[^>]*>(.*?)</row>', sheet_xml, re.S)
    data = [(int(number), body) for number, body in rows if int(number) > 1]
    body = data[-1][1] if data else ""
    attrs = {}
    for match in re.finditer(r'<c r="([A-Z]+)\d+"([^>]*?)(?:/>|>)', body):
        attrs[match.group(1)] = re.sub(r'\s+t="[^"]*"', "", match.group(2))
    return attrs


def _cell_xml(column, row_number, value, attrs):
    base = attrs.get(column, "")
    if not base.startswith(" ") and base:
        base = " " + base
    ref = f"{column}{row_number}"
    if column in NUMERIC_COLS:
        if value in (None, ""):
            return f'<c r="{ref}"{base} t="n"/>'
        return f'<c r="{ref}"{base} t="n"><v>{xml_number(value)}</v></c>'
    return f'<c r="{ref}"{base} t="inlineStr"><is><t>{xml_text(value)}</t></is></c>'


def _row_xml(row_number, values, attrs):
    letters = "ABCDEFGHIJKLM"
    cells = "".join(_cell_xml(letter, row_number, values[index], attrs) for index, letter in enumerate(letters))
    return f'<row r="{row_number}" ht="16" customHeight="1">{cells}</row>'


def _column_letter(index):
    letters = ""
    number = index + 1
    while number:
        number, remainder = divmod(number - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def xml_cell_text(value, limit=8000):
    text = str(value or "").replace("\r\n", "\n").replace("\r", "\n")
    text = "".join(ch for ch in text if ch in "\n\t" or ord(ch) >= 32)[:limit]
    if text[:1] in "=+-@":
        text = "'" + text
    return html.escape(text, quote=False)


def _table_cell(column, row_number, value):
    ref = f"{_column_letter(column)}{row_number}"
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return f'<c r="{ref}" t="n"><v>{xml_number(value)}</v></c>'
    return f'<c r="{ref}" t="inlineStr"><is><t>{xml_text(value)}</t></is></c>'


def _sheet_title(name):
    cleaned = "".join(ch for ch in one_line(name, 31) if ch not in "[]:*?/\\")
    return cleaned or "Sayfa"


def _export_style(kind, alt):
    if kind == "money":
        return 5 if alt else 4
    if kind == "date":
        return 7 if alt else 6
    return 3 if alt else 2


def _export_width(header, values, kind):
    if kind == "long":
        return 46
    if kind == "date":
        return 16
    if kind == "money":
        return 16
    longest = len(str(header))
    for value in values[:100]:
        for line in str(value or "").split("\n"):
            longest = max(longest, min(len(line), 42))
    return max(14, min(36, longest + 3))


def _export_cell(column, row_number, value, kind, style):
    ref = f"{_column_letter(column)}{row_number}"
    if kind == "money" and isinstance(value, (int, float)) and not isinstance(value, bool):
        return f'<c r="{ref}" s="{style}"><v>{xml_number(value)}</v></c>'
    if kind == "date" and isinstance(value, date):
        return f'<c r="{ref}" s="{style}"><v>{excel_serial(value)}</v></c>'
    text = xml_cell_text(value)
    space = ' xml:space="preserve"' if "\n" in str(value or "") else ""
    return f'<c r="{ref}" s="{style}" t="inlineStr"><is><t{space}>{text}</t></is></c>'


def _worksheet_xml(headers, data_rows, kinds):
    kinds = list(kinds or [])
    while len(kinds) < len(headers):
        kinds.append("text")
    columns = []
    for index, header in enumerate(headers):
        samples = [row[index] if index < len(row) else "" for row in data_rows]
        width = _export_width(header, samples, kinds[index])
        columns.append(f'<col min="{index + 1}" max="{index + 1}" width="{width}" customWidth="1"/>')
    lines = []
    header_cells = "".join(_export_cell(index, 1, title, "text", 1) for index, title in enumerate(headers))
    lines.append(f'<row r="1" ht="24" customHeight="1">{header_cells}</row>')
    for offset, values in enumerate(data_rows):
        row_number = offset + 2
        alt = offset % 2 == 1
        lines_needed = 1
        for index, kind in enumerate(kinds):
            if kind != "long" or index >= len(values):
                continue
            text = str(values[index] or "")
            lines_needed = max(lines_needed, text.count("\n") + 1, min(5, (len(text) // 46) + 1))
        height = min(90, 18 * lines_needed)
        cells = []
        for index, header in enumerate(headers):
            value = values[index] if index < len(values) else ""
            cells.append(_export_cell(index, row_number, value, kinds[index], _export_style(kinds[index], alt)))
        lines.append(f'<row r="{row_number}" ht="{height}" customHeight="1">{"".join(cells)}</row>')
    last_column = _column_letter(max(len(headers) - 1, 0))
    last_row = max(len(data_rows) + 1, 1)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>'
        f'<dimension ref="A1:{last_column}{last_row}"/>'
        '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>'
        '<sheetFormatPr defaultRowHeight="18"/>'
        f'<cols>{"".join(columns)}</cols>'
        f'<sheetData>{"".join(lines)}</sheetData>'
        f'<autoFilter ref="A1:{last_column}{last_row}"/>'
        '<pageSetup orientation="landscape" fitToWidth="1" fitToHeight="0" paperSize="9"/>'
        '</worksheet>'
    )


_EXPORT_STYLES = """<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="2">
<numFmt numFmtId="164" formatCode="#,##0.00"/>
<numFmt numFmtId="165" formatCode="dd.mm.yyyy"/>
</numFmts>
<fonts count="2">
<font><sz val="11"/><name val="Calibri"/></font>
<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
</fonts>
<fills count="4">
<fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF1B2836"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFF4F7FA"/></patternFill></fill>
</fills>
<borders count="2">
<border/>
<border>
<left style="thin"><color rgb="FFD5DEE8"/></left>
<right style="thin"><color rgb="FFD5DEE8"/></right>
<top style="thin"><color rgb="FFD5DEE8"/></top>
<bottom style="thin"><color rgb="FFD5DEE8"/></bottom>
</border>
</borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="8">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
<xf numFmtId="0" fontId="0" fillId="3" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
<xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf>
<xf numFmtId="164" fontId="0" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf>
<xf numFmtId="165" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf>
<xf numFmtId="165" fontId="0" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf>
</cellXfs>
</styleSheet>"""


def write_workbook(path, sheets):
    path = Path(path)
    if path.suffix.lower() != ".xlsx":
        raise ValueError("Dosya .xlsx olmalı.")
    if not sheets:
        raise ValueError("Yazılacak sayfa yok.")
    used = set()
    prepared = []
    for index, sheet in enumerate(sheets, start=1):
        title = _sheet_title(sheet["name"])
        base = title
        suffix = 2
        while title in used:
            title = _sheet_title(f"{base[:28]} {suffix}")
            suffix += 1
        used.add(title)
        prepared.append((index, title, _worksheet_xml(sheet["headers"], sheet["rows"], sheet.get("kinds"))))
    overrides = [
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
    ]
    overrides.extend(
        f'<Override PartName="/xl/worksheets/sheet{index}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        for index, _title, _xml in prepared
    )
    sheet_tags = "".join(
        f'<sheet name="{html.escape(title, quote=True)}" sheetId="{index}" r:id="rId{index}"/>'
        for index, title, _xml in prepared
    )
    relationships = [
        f'<Relationship Id="rId{index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{index}.xml"/>'
        for index, _title, _xml in prepared
    ]
    relationships.append(
        f'<Relationship Id="rId{len(prepared) + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    )
    files = {
        "[Content_Types].xml": (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            + "".join(overrides)
            + "</Types>"
        ),
        "_rels/.rels": """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>""",
        "xl/workbook.xml": (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            f"<sheets>{sheet_tags}</sheets></workbook>"
        ),
        "xl/_rels/workbook.xml.rels": (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            + "".join(relationships)
            + "</Relationships>"
        ),
        "xl/styles.xml": _EXPORT_STYLES,
    }
    for index, _title, xml in prepared:
        files[f"xl/worksheets/sheet{index}.xml"] = xml
    temporary = path.with_suffix(".writing.xlsx")
    try:
        with zipfile.ZipFile(temporary, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            for name, payload in files.items():
                archive.writestr(name, payload)
        with zipfile.ZipFile(temporary) as archive:
            for index, _title, _xml in prepared:
                ET.fromstring(archive.read(f"xl/worksheets/sheet{index}.xml"))
            ET.fromstring(archive.read("xl/styles.xml"))
        os.replace(temporary, path)
    except Exception:
        temporary.unlink(missing_ok=True)
        raise


def write_simple_xlsx(path, sheet_name, headers, data_rows):
    write_workbook(path, [{"name": sheet_name, "headers": headers, "rows": data_rows, "kinds": ["text"] * len(headers)}])


def _cell(row, key):
    if key not in row.keys() or row[key] is None:
        return ""
    return str(row[key]).strip()


def _money_cell(row, key):
    if key not in row.keys() or row[key] in (None, ""):
        return ""
    return float(row[key])


def _date_cell(row, key):
    return parse_day(_cell(row, key)) or ""


def _task_note(row):
    raw = row["note_json"] if "note_json" in row.keys() else ""
    return note_plain(raw).strip()


def _invoice_source(row):
    return {"klasor": "Klasör", "eposta": "E-posta"}.get(row["source"] or "", "Elle")


def _invoice_file(row):
    name = _cell(row, "file_name")
    if name:
        return name
    attachment = _cell(row, "attachment")
    return Path(attachment).name if attachment else ""


def _ledger_docs(row):
    raw = row["docs_json"] if "docs_json" in row.keys() else ""
    return "\n".join(item["name"] for item in ledger_documents(raw))


def _export_sheet(name, columns, records):
    return {
        "name": name,
        "headers": [header for header, _kind, _getter in columns],
        "kinds": [kind for _header, kind, _getter in columns],
        "rows": [[getter(record) for _header, _kind, getter in columns] for record in records],
    }


_TASK_EXPORT = [
    ("Anahtar", "text", lambda row: ""),
    ("Tür", "text", lambda row: _cell(row, "issue_type")),
    ("Görev", "text", lambda row: _cell(row, "title")),
    ("Durum", "text", lambda row: _cell(row, "status")),
    ("Öncelik", "text", lambda row: _cell(row, "priority")),
    ("Sorumlu", "text", lambda row: _cell(row, "assignee")),
    ("Başlangıç", "date", lambda row: _date_cell(row, "start_date")),
    ("Son tarih", "date", lambda row: _date_cell(row, "due_date")),
    ("Tekrar", "text", lambda row: _cell(row, "repeat")),
    ("Etiket", "text", lambda row: _cell(row, "tag")),
    ("Açıklama", "long", lambda row: _cell(row, "description")),
    ("Yorum", "long", _task_note),
]
_REMINDER_EXPORT = [
    ("Başlık", "text", lambda row: _cell(row, "title")),
    ("Durum", "text", lambda row: _cell(row, "status")),
    ("Tarih", "date", lambda row: _date_cell(row, "remind_date")),
    ("Saat", "text", lambda row: _cell(row, "remind_time")),
    ("Tekrar", "text", lambda row: _cell(row, "repeat")),
    ("Açıklama", "long", lambda row: _cell(row, "description")),
]
_CONTACT_EXPORT = [
    ("Kayıt", "text", lambda row: _cell(row, "record_kind")),
    ("Kapsam", "text", lambda row: _cell(row, "scope")),
    ("Ad", "text", lambda row: _cell(row, "name")),
    ("Yetkili", "text", lambda row: _cell(row, "contact_person")),
    ("Telefon", "text", lambda row: _cell(row, "phone")),
    ("Dahili", "text", lambda row: _cell(row, "extension")),
    ("E-posta", "text", lambda row: _cell(row, "email")),
    ("Vergi / kimlik no", "text", lambda row: _cell(row, "tax_no")),
    ("Adres", "long", lambda row: _cell(row, "address")),
    ("Notlar", "long", lambda row: _cell(row, "notes")),
]
_LEDGER_EXPORT = [
    ("Tür", "text", lambda row: _cell(row, "kind")),
    ("Kategori", "text", lambda row: _cell(row, "category")),
    ("Tarih", "date", lambda row: _date_cell(row, "trans_date")),
    ("Tutar", "money", lambda row: _money_cell(row, "amount")),
    ("Ödeme", "text", lambda row: _cell(row, "payment_method")),
    ("Firma / kişi", "text", lambda row: _cell(row, "party")),
    ("Açıklama", "long", lambda row: _cell(row, "description")),
    ("Belgeler", "long", _ledger_docs),
]
_INVOICE_EXPORT = [
    ("Tür", "text", lambda row: _cell(row, "kind")),
    ("Durum", "text", lambda row: _cell(row, "status")),
    ("Firma / kişi", "text", lambda row: _cell(row, "party")),
    ("Fatura no", "text", lambda row: _cell(row, "invoice_no")),
    ("Fatura tarihi", "date", lambda row: _date_cell(row, "invoice_date")),
    ("Vade", "date", lambda row: _date_cell(row, "due_date")),
    ("Tutar", "money", lambda row: _money_cell(row, "amount")),
    ("Para birimi", "text", lambda row: _cell(row, "currency") or "TRY"),
    ("Açıklama", "long", lambda row: _cell(row, "description")),
    ("Vergi no", "text", lambda row: _cell(row, "tax_no")),
    ("Kaynak", "text", _invoice_source),
    ("Dosya", "text", _invoice_file),
    ("Excel durumu", "text", lambda row: _cell(row, "excel_status")),
    ("Rapor notu", "long", lambda row: _cell(row, "report_reason")),
]
_PROJECT_EXPORT = [
    ("Proje", "text", lambda row: _cell(row, "name")),
    ("Anahtar", "text", lambda row: _cell(row, "proj_key")),
    ("Durum", "text", lambda row: _cell(row, "status")),
    ("Başlangıç", "date", lambda row: _date_cell(row, "start_date")),
    ("Bitiş", "date", lambda row: _date_cell(row, "end_date")),
    ("Bütçe", "money", lambda row: _money_cell(row, "budget")),
    ("Teklif sayısı", "text", lambda row: ""),
    ("Açıklama", "long", lambda row: _cell(row, "description")),
    ("Notlar", "long", lambda row: _cell(row, "notes")),
]
_OFFER_EXPORT = [
    ("Proje", "text", lambda pair: _cell(pair[0], "name")),
    ("Firma", "text", lambda pair: _cell(pair[1], "vendor")),
    ("Başlık", "text", lambda pair: _cell(pair[1], "title")),
    ("Tutar", "money", lambda pair: _money_cell(pair[1], "amount")),
    ("Tarih", "date", lambda pair: _date_cell(pair[1], "offer_date")),
    ("Artılar", "long", lambda pair: _cell(pair[1], "pros")),
    ("Eksiler", "long", lambda pair: _cell(pair[1], "cons")),
    ("Not", "long", lambda pair: _cell(pair[1], "notes")),
    ("Dosya", "text", lambda pair: Path(pair[1]["file_path"]).name if pair[1]["file_path"] else ""),
]


def create_invoice_workbook(path):
    path = Path(path)
    header = "".join(
        f'<c r="{letter}1" t="inlineStr"><is><t>{xml_text(title)}</t></is></c>'
        for letter, title in zip("ABCDEFGHIJKLM", EXCEL_HEADERS)
    )
    sheet = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<sheetData>'
        f'<row r="1">{header}</row>'
        "</sheetData></worksheet>"
    )
    files = {
        "[Content_Types].xml": """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>""",
        "_rels/.rels": """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>""",
        "xl/workbook.xml": """<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Faturalar" sheetId="1" r:id="rId1"/></sheets></workbook>""",
        "xl/_rels/workbook.xml.rels": """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>""",
        "xl/styles.xml": """<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>""",
        "xl/worksheets/sheet1.xml": sheet,
    }
    _write_xlsx(path, files)


def _write_xlsx(path, files):
    path = Path(path)
    temporary = path.with_suffix(".writing.xlsx")
    names = list(files)
    if "[Content_Types].xml" in names:
        names.remove("[Content_Types].xml")
        names.insert(0, "[Content_Types].xml")
    try:
        with zipfile.ZipFile(temporary, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            for name in names:
                archive.writestr(name, files[name])
        with zipfile.ZipFile(temporary) as archive:
            ET.fromstring(archive.read(_sheet_target({name: archive.read(name) for name in archive.namelist()})))
        os.replace(temporary, path)
    except Exception:
        temporary.unlink(missing_ok=True)
        raise


def append_excel_invoices(path, new_rows):
    path = Path(path)
    if not new_rows:
        return 0
    if not path.exists():
        if path.suffix.lower() != ".xlsx" or not acceptable_watch_dir(path.parent):
            raise ValueError("Excel yalnızca seçilen fatura klasörüne yazılır.")
        create_invoice_workbook(path)
    if path.suffix.lower() != ".xlsx" or path.is_symlink():
        raise ValueError("Excel dosyası .xlsx olmalı.")
    with zipfile.ZipFile(path) as workbook:
        files = {name: workbook.read(name) for name in workbook.namelist()}
    sheet_name = _sheet_target(files)
    sheet = files[sheet_name].decode("utf-8")
    if "</sheetData>" not in sheet:
        raise ValueError("Excel sayfası beklenen yapıda değil.")
    attrs = _column_attrs(sheet)
    numbers = [int(value) for value in re.findall(r'<row r="(\d+)"', sheet)]
    last = max(numbers) if numbers else 1
    existing = {path_key(row["path"]) for row in read_excel_invoices(path)}
    built = []
    for row in new_rows:
        if path_key(row[1]) in existing:
            continue
        last += 1
        values = [
            row[0],
            row[1],
            row[2],
            row[3],
            excel_serial(row[4]) if isinstance(row[4], date) else row[4],
            row[5],
            row[6],
            row[7],
            row[8],
            row[9] or "TRY",
            row[10],
            row[11],
            excel_serial(row[12]) if isinstance(row[12], date) else row[12],
        ]
        built.append(_row_xml(last, values, attrs))
        existing.add(path_key(row[1]))
    if not built:
        return 0
    backup = path.with_name(path.stem + ".onceki.xlsx")
    if not backup.exists():
        shutil.copy2(path, backup)
    sheet = sheet.replace("</sheetData>", "".join(built) + "</sheetData>", 1)
    sheet = re.sub(r'(<dimension[^>]*ref=")[^"]+(")', rf"\g<1>A1:M{last}\2", sheet, count=1)
    files[sheet_name] = sheet.encode("utf-8")
    table_name = _table_target(files, sheet_name)
    if table_name and table_name in files:
        table = files[table_name].decode("utf-8")
        table = re.sub(r'ref="A1:[A-Z]+\d+"', f'ref="A1:M{last}"', table)
        files[table_name] = table.encode("utf-8")
    _write_xlsx(path, files)
    return len(built)


def valid_host(host):
    return bool(re.fullmatch(r"[A-Za-z0-9.-]{1,253}", host)) and ".." not in host and not host.startswith("-")


def valid_user(user):
    return bool(re.fullmatch(r"[^\s\x00-\x1f]{1,200}", user or ""))


def valid_password(password):
    return bool(password) and len(password) <= 200 and "\x00" not in password and "\n" not in password and "\r" not in password


def valid_mailbox(name):
    return bool(re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9 _./-]{0,79}", name or ""))


MAIL_PROGRAMS = {
    "Mail": (LOCAL_MAIL_HOST, "0", "INBOX", "Mac’teki Mail uygulamasının gelen kutusu okunur. Adres ve parola yazılmaz. İlk seferde Mail’e erişim izni istenir."),
    "Gmail": ("imap.gmail.com", "993", "INBOX", "Gmail normal parolayı kabul etmez. Google hesabından uygulama parolası oluştur."),
    "Outlook": ("outlook.office365.com", "993", "INBOX", "Outlook veya Hotmail adresini ve parolasını yaz."),
    "iCloud": ("imap.mail.me.com", "993", "INBOX", "iCloud için Apple kimliğinden uygulama parolası gerekir."),
    "Yandex": ("imap.yandex.com", "993", "INBOX", "Yandex adresini ve parolasını yaz."),
    "Diğer": ("", "993", "INBOX", "Kurumsal posta için sunucu yazman yeter. Port boşsa güvenli bağlantı kullanılır."),
}


def _win_credential():
    class Credential(ctypes.Structure):
        _fields_ = [
            ("Flags", ctypes.c_uint32),
            ("Type", ctypes.c_uint32),
            ("TargetName", ctypes.c_wchar_p),
            ("Comment", ctypes.c_wchar_p),
            ("LastWritten", ctypes.c_uint64),
            ("CredentialBlobSize", ctypes.c_uint32),
            ("CredentialBlob", ctypes.c_void_p),
            ("Persist", ctypes.c_uint32),
            ("AttributeCount", ctypes.c_uint32),
            ("Attributes", ctypes.c_void_p),
            ("TargetAlias", ctypes.c_wchar_p),
            ("UserName", ctypes.c_wchar_p),
        ]
    return Credential, ctypes.windll.advapi32


def keychain_set(account, password):
    if sys.platform == "win32":
        credential, advapi = _win_credential()
        blob = ctypes.create_string_buffer(password.encode("utf-8"))
        item = credential()
        item.Type = 1
        item.TargetName = f"{KEYCHAIN_SERVICE}:{account}"
        item.UserName = account
        item.Comment = APP_NAME
        item.CredentialBlobSize = len(blob.raw) - 1
        item.CredentialBlob = ctypes.cast(blob, ctypes.c_void_p)
        item.Persist = 2
        if not advapi.CredWriteW(ctypes.byref(item), 0):
            raise ValueError("Parola Windows Kimlik Yöneticisi'ne yazılamadı.")
        return
    result = subprocess.run(
        ["security", "add-generic-password", "-U", "-s", KEYCHAIN_SERVICE, "-a", account, "-w", password],
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        raise ValueError("Parola Anahtar Zinciri'ne yazılamadı.")


def keychain_get(account):
    if sys.platform == "win32":
        credential, advapi = _win_credential()
        holder = ctypes.c_void_p()
        if not advapi.CredReadW(f"{KEYCHAIN_SERVICE}:{account}", 1, 0, ctypes.byref(holder)):
            return ""
        try:
            item = ctypes.cast(holder, ctypes.POINTER(credential)).contents
            raw = ctypes.string_at(item.CredentialBlob, item.CredentialBlobSize)
            return raw.decode("utf-8")
        finally:
            advapi.CredFree(holder)
    result = subprocess.run(
        ["security", "find-generic-password", "-s", KEYCHAIN_SERVICE, "-a", account, "-w"],
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        return ""
    return result.stdout.rstrip("\n")


def header_text(value):
    try:
        return one_line(str(make_header(decode_header(value or ""))), 180)
    except (ValueError, TypeError, email.errors.HeaderParseError):
        return one_line(value, 180)


def fetch_mail_invoices(host, port, user, password, mailbox):
    if not valid_host(host) or not valid_user(user) or not valid_password(password) or not valid_mailbox(mailbox):
        raise ValueError("E-posta ayarları geçersiz.")
    if not isinstance(port, int) or not 1 <= port <= 65535:
        raise ValueError("Port geçersiz.")
    saved = []
    client = imaplib.IMAP4_SSL(host, port, timeout=20)
    try:
        client.login(user, password)
        status, _ = client.select(mailbox, readonly=True)
        if status != "OK":
            raise ValueError("Posta klasörü açılamadı.")
        since = (date.today() - timedelta(days=MAIL_WINDOW_DAYS)).strftime("%d-%b-%Y")
        status, data = client.uid("SEARCH", None, "SINCE", since)
        if status != "OK" or not data or not data[0]:
            return []
        uids = data[0].split()[-MAIL_MESSAGE_CAP:]
        for uid in uids:
            status, fetched = client.uid("FETCH", uid, "(BODY.PEEK[])")
            if status != "OK":
                continue
            raw = b""
            for item in fetched:
                if isinstance(item, tuple) and len(item) > 1:
                    raw = item[1]
            if not raw:
                continue
            message = email.message_from_bytes(raw, policy=email.policy.default)
            sender = header_text(message.get("From"))
            party = re.sub(r"<[^>]+>", "", sender).strip(' "') or "E-posta"
            subject = header_text(message.get("Subject")) or "E-posta eki"
            uid_text = uid.decode("ascii", "ignore")
            for part in message.walk():
                if part.get_content_maintype() == "multipart":
                    continue
                filename = part.get_filename()
                if not filename and part.get_content_type() == "application/pdf":
                    filename = f"fatura-{uid_text}.pdf"
                if not filename:
                    continue
                filename = safe_name(filename)
                ext = Path(filename).suffix.lower()
                if ext not in ALLOWED_EXT:
                    continue
                payload = part.get_payload(decode=True)
                if not payload or len(payload) > MAX_BYTES or not signature_ok(ext, payload[:16]):
                    continue
                ATTACH_DIR.mkdir(parents=True, exist_ok=True)
                os.chmod(ATTACH_DIR, 0o700)
                index = 1 + sum(1 for item in saved if item["source_key"].startswith(f"mail:{uid_text}:"))
                dest = ATTACH_DIR / f"mail-{uid_text}-{index}-{filename}"
                if not dest.exists():
                    dest.write_bytes(payload)
                    os.chmod(dest, 0o600)
                if not file_signature_ok(dest):
                    dest.unlink(missing_ok=True)
                    continue
                saved.append(
                    {
                        "source_key": f"mail:{uid_text}:{filename}",
                        "party": party[:120],
                        "description": subject,
                        "attachment": str(dest),
                        "filename": filename,
                    }
                )
    finally:
        try:
            client.logout()
        except imaplib.IMAP4.error:
            pass
    return saved


MAIL_APP_SCRIPT = r'''on run argv
	set destDir to item 1 of argv
	set manifestPath to item 2 of argv
	set dayCount to (item 3 of argv) as integer
	set capCount to (item 4 of argv) as integer
	set maxBytes to (item 5 of argv) as integer
	set allowed to {".pdf", ".png", ".jpg", ".jpeg", ".webp", ".tif", ".tiff", ".heic"}
	tell application "Mail"
		set cutoff to (current date) - (dayCount * days)
		set checked to 0
		repeat with acct in accounts
			if checked ≥ capCount then exit repeat
			repeat with mb in mailboxes of acct
				if checked ≥ capCount then exit repeat
				set mbName to name of mb
				if mbName is "INBOX" or mbName is "Inbox" or mbName is "Gelen Kutusu" then
					set total to count of messages of mb
					set span to capCount - checked
					if span > total then set span to total
					if span > 0 then
					-- Bulk reads: per-message Apple events take ~0.25s each on a large inbox.
					set idList to id of messages 1 thru span of mb
					set dateList to date received of messages 1 thru span of mb
					set attLists to name of every mail attachment of messages 1 thru span of mb
					repeat with i from 1 to span
						try
							if (item i of dateList) comes before cutoff then exit repeat
						end try
						set checked to checked + 1
						set msg to missing value
						if (count of (item i of attLists)) > 0 then
							set msg to message i of mb
							if (id of msg) is not (item i of idList) then
								set msg to missing value
								try
									set msg to first message of mb whose id is (item i of idList)
								end try
							end if
						end if
						if msg is not missing value then
			set msgId to (item i of idList) as string
			set snd to ""
			set subj to ""
			try
				set snd to sender of msg
			end try
			try
				set subj to subject of msg
			end try
			if snd is missing value then set snd to ""
			if subj is missing value then set subj to ""
			set idx to 0
			repeat with att in (mail attachments of msg)
				set idx to idx + 1
				set attName to ""
				try
					set attName to name of att
				end try
				if attName is missing value or attName is "" then set attName to "ek-" & idx
				set mimeType to ""
				try
					set mimeType to MIME type of att
				end try
				set AppleScript's text item delimiters to "."
				set nameParts to text items of attName
				set ext to ""
				if (count of nameParts) > 1 then set ext to "." & last item of nameParts
				set AppleScript's text item delimiters to ""
				try
					set ext to do shell script "printf %s " & quoted form of ext & " | /usr/bin/tr A-Z a-z"
				end try
				if ext is "" then
					if mimeType is "application/pdf" then set ext to ".pdf"
					if mimeType is "image/png" then set ext to ".png"
					if mimeType is "image/jpeg" then set ext to ".jpg"
					if mimeType is "image/webp" then set ext to ".webp"
					if mimeType is "image/tiff" then set ext to ".tif"
					if mimeType is "image/heic" then set ext to ".heic"
				end if
				if ext is in allowed then
					set tooBig to false
					try
						if (file size of att) > maxBytes then set tooBig to true
					end try
					if tooBig is false then
						try
							set stored to do shell script "printf %s " & quoted form of (msgId & "-" & idx & ext) & " | /usr/bin/tr -c 'A-Za-z0-9._-' '_'"
							set targetPosix to destDir & "/" & stored
							save att in (POSIX file targetPosix)
							set payload to msgId & tab & snd & tab & subj & tab & attName & tab & stored
							set encoded to do shell script "printf %s " & quoted form of payload & " | /usr/bin/openssl base64 -A"
							do shell script "printf '%s\\n' " & quoted form of encoded & " >> " & quoted form of manifestPath
						end try
					end if
				end if
			end repeat
						end if
					end repeat
					end if
				end if
			end repeat
		end repeat
	end tell
	return "ok"
end run
'''


def mail_app_running():
    try:
        result = subprocess.run(
            ["osascript", "-e", 'application "Mail" is running'],
            capture_output=True,
            text=True,
            timeout=8,
        )
    except (OSError, subprocess.TimeoutExpired):
        return False
    return result.returncode == 0 and result.stdout.strip() == "true"


def _mail_app_failure(stderr):
    text = stderr or ""
    if "-1743" in text or "not authorized" in text.lower() or "not permitted" in text.lower():
        return "Mail’e erişim izni yok. Sistem Ayarları, Gizlilik ve Güvenlik, Otomasyon bölümünden Python’un Mail’i kullanmasına izin ver."
    if "-1728" in text or "inbox" in text.lower():
        return "Mail’de gelen kutusu açılamadı."
    return "Mac Mail okunamadı."


def _party_from_sender(sender):
    party = re.sub(r"<[^>]+>", "", sender or "").strip(' "')
    return party[:120] or "E-posta"


def _collect_mail_app_files(tmp_path, manifest):
    saved = []
    if not manifest.is_file():
        return saved
    root = tmp_path.resolve()
    for line in manifest.read_text(encoding="utf-8", errors="replace").splitlines():
        try:
            raw = base64.b64decode(line.strip(), validate=True)
            msg_id, sender, subject, filename, stored = raw.decode("utf-8", "replace").split("\t", 4)
        except (ValueError, TypeError):
            continue
        src = tmp_path / Path(stored).name
        if not src.is_file() or src.is_symlink() or not is_within(src, root):
            continue
        filename = safe_name(filename)
        ext = Path(filename).suffix.lower()
        if ext not in ALLOWED_EXT:
            continue
        try:
            payload = src.read_bytes()
        except OSError:
            continue
        if not payload or len(payload) > MAX_BYTES or not signature_ok(ext, payload[:16]):
            continue
        uid_text = re.sub(r"[^0-9]", "", msg_id)[:12] or "0"
        ATTACH_DIR.mkdir(parents=True, exist_ok=True)
        os.chmod(ATTACH_DIR, 0o700)
        index = 1 + sum(1 for item in saved if item["source_key"].startswith(f"mailapp:{uid_text}:"))
        dest = ATTACH_DIR / f"mail-{uid_text}-{index}-{filename}"
        if not dest.exists():
            dest.write_bytes(payload)
            os.chmod(dest, 0o600)
        if not file_signature_ok(dest):
            dest.unlink(missing_ok=True)
            continue
        saved.append(
            {
                "source_key": f"mailapp:{uid_text}:{index}:{filename}",
                "party": _party_from_sender(sender),
                "description": one_line(subject, 180) or "E-posta eki",
                "attachment": str(dest),
                "filename": filename,
            }
        )
    return saved


def fetch_mail_app_invoices(allow_launch=True):
    if sys.platform != "darwin":
        raise ValueError("Mac’teki Mail yalnız Mac bilgisayarda okunur.")
    if not allow_launch and not mail_app_running():
        return None
    # Mail is sandboxed and can only save attachments under ~/Downloads, not $TMPDIR.
    downloads = Path.home() / "Downloads"
    downloads.mkdir(exist_ok=True)
    with tempfile.TemporaryDirectory(prefix=".vkv-mail-", dir=downloads) as tmp:
        tmp_path = Path(tmp)
        os.chmod(tmp_path, 0o700)
        manifest = tmp_path / "liste.txt"
        scan_cap = 120 if allow_launch else 40
        try:
            result = subprocess.run(
                [
                    "osascript",
                    "-",
                    str(tmp_path),
                    str(manifest),
                    str(MAIL_WINDOW_DAYS),
                    str(scan_cap),
                    str(MAX_BYTES),
                ],
                input=MAIL_APP_SCRIPT,
                capture_output=True,
                text=True,
                timeout=240 if allow_launch else 120,
            )
        except subprocess.TimeoutExpired as exc:
            raise ValueError("Mail yanıt vermedi.") from exc
        except OSError as exc:
            raise ValueError("Mac Mail okunamadı.") from exc
        if result.returncode != 0:
            raise ValueError(_mail_app_failure(result.stderr))
        return _collect_mail_app_files(tmp_path, manifest)


def render_preview(src, size):
    src = Path(src)
    stamp = src.stat().st_mtime_ns
    digest = hashlib.sha256(f"{src}|{stamp}|{size}".encode()).hexdigest()[:24]
    folder = PREVIEW_DIR / digest
    folder.mkdir(parents=True, exist_ok=True)
    os.chmod(PREVIEW_DIR, 0o700)
    existing = list(folder.glob("*.png"))
    if existing:
        return existing[0]
    if sys.platform == "win32":
        if src.suffix.lower() in {".png", ".gif"}:
            return src
        return None
    subprocess.run(
        ["qlmanage", "-t", "-s", str(int(size)), "-o", str(folder), str(src)],
        check=False,
        timeout=25,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
    images = list(folder.glob("*.png"))
    return images[0] if images else None


def fit_photo(path, max_width, max_height):
    image = tk.PhotoImage(file=str(path))
    scale = 1
    while image.width() // scale > max_width or image.height() // scale > max_height:
        scale += 1
        if scale > 12:
            break
    if scale > 1:
        image = image.subsample(scale, scale)
    return image


class Database:
    def __init__(self):
        DATA_DIR.mkdir(parents=True, exist_ok=True)
        ATTACH_DIR.mkdir(parents=True, exist_ok=True)
        LEDGER_DIR.mkdir(parents=True, exist_ok=True)
        PREVIEW_DIR.mkdir(parents=True, exist_ok=True)
        NOTE_DIR.mkdir(parents=True, exist_ok=True)
        for folder in (DATA_DIR, ATTACH_DIR, LEDGER_DIR, PREVIEW_DIR, NOTE_DIR):
            try:
                os.chmod(folder, 0o700)
            except OSError:
                if sys.platform != "win32":
                    raise
        if sys.platform == "win32" and DB_PATH.is_file():
            try:
                os.chmod(DB_PATH, 0o600)
            except OSError:
                pass
        self.conn = sqlite3.connect(DB_PATH, timeout=10)
        self.conn.row_factory = sqlite3.Row
        self.conn.execute("PRAGMA foreign_keys=ON")
        self.conn.execute("PRAGMA journal_mode=WAL")
        self.init_db()
        try:
            os.chmod(DB_PATH, 0o600)
        except OSError:
            if sys.platform != "win32":
                raise

    def init_db(self):
        self.conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS tasks (
              id INTEGER PRIMARY KEY, title TEXT NOT NULL, description TEXT, assignee TEXT,
              priority TEXT, status TEXT, start_date TEXT, due_date TEXT, repeat TEXT, tag TEXT,
              created_at TEXT DEFAULT CURRENT_TIMESTAMP);
            CREATE TABLE IF NOT EXISTS invoices (
              id INTEGER PRIMARY KEY, kind TEXT, party TEXT NOT NULL, invoice_no TEXT,
              description TEXT, amount REAL, invoice_date TEXT, due_date TEXT, status TEXT,
              attachment TEXT, created_at TEXT DEFAULT CURRENT_TIMESTAMP);
            CREATE TABLE IF NOT EXISTS reminders (
              id INTEGER PRIMARY KEY, title TEXT NOT NULL, description TEXT, remind_date TEXT,
              remind_time TEXT, repeat TEXT, status TEXT, last_notified TEXT,
              created_at TEXT DEFAULT CURRENT_TIMESTAMP);
            CREATE TABLE IF NOT EXISTS contacts (
              id INTEGER PRIMARY KEY, name TEXT NOT NULL, contact_person TEXT, phone TEXT,
              email TEXT, tax_no TEXT, address TEXT, notes TEXT, created_at TEXT DEFAULT CURRENT_TIMESTAMP);
            CREATE TABLE IF NOT EXISTS transactions (
              id INTEGER PRIMARY KEY, kind TEXT, category TEXT, description TEXT, amount REAL,
              trans_date TEXT, party TEXT, payment_method TEXT, created_at TEXT DEFAULT CURRENT_TIMESTAMP);
            CREATE TABLE IF NOT EXISTS projects (
              id INTEGER PRIMARY KEY, name TEXT NOT NULL, description TEXT, start_date TEXT,
              end_date TEXT, status TEXT, budget REAL, notes TEXT, created_at TEXT DEFAULT CURRENT_TIMESTAMP);
            CREATE TABLE IF NOT EXISTS notify_log (key TEXT PRIMARY KEY, last_on TEXT);
            CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT);
            CREATE TABLE IF NOT EXISTS ignored_sources (key TEXT PRIMARY KEY);
            CREATE TABLE IF NOT EXISTS reports (
              id INTEGER PRIMARY KEY, invoice_id INTEGER, title TEXT, invoice_no TEXT,
              amount REAL, reason TEXT, saved_path TEXT, created_at TEXT DEFAULT CURRENT_TIMESTAMP);
            CREATE TABLE IF NOT EXISTS offers (
              id INTEGER PRIMARY KEY, project_id INTEGER NOT NULL, vendor TEXT NOT NULL,
              title TEXT, amount REAL, offer_date TEXT, file_path TEXT, pros TEXT, cons TEXT,
              notes TEXT, source_text TEXT, created_at TEXT DEFAULT CURRENT_TIMESTAMP);
            """
        )
        self._add_column("invoices", "source", "TEXT")
        self._add_column("invoices", "source_key", "TEXT")
        self._add_column("invoices", "currency", "TEXT")
        self._add_column("invoices", "tax_no", "TEXT")
        self._add_column("invoices", "excel_status", "TEXT")
        self._add_column("invoices", "file_name", "TEXT")
        self._add_column("invoices", "report_reason", "TEXT")
        self._add_column("invoices", "report_files", "TEXT")
        self._add_column("invoices", "paid_amount", "REAL")
        self._add_column("invoices", "events_json", "TEXT")
        self._add_column("contacts", "record_kind", "TEXT")
        self._add_column("contacts", "scope", "TEXT")
        self._add_column("contacts", "extension", "TEXT")
        self._add_column("offers", "source_text", "TEXT")
        self._add_column("offers", "url", "TEXT")
        self._add_column("tasks", "waiting_since", "TEXT")
        self._add_column("tasks", "sla_next", "TEXT")
        self._add_column("tasks", "note_json", "TEXT")
        self._add_column("tasks", "project_id", "INTEGER")
        self._add_column("tasks", "issue_no", "INTEGER")
        self._add_column("tasks", "issue_type", "TEXT")
        self._add_column("projects", "proj_key", "TEXT")
        self._add_column("transactions", "docs_json", "TEXT")
        self.conn.execute(
            "CREATE UNIQUE INDEX IF NOT EXISTS idx_invoices_source ON invoices(source_key) "
            "WHERE source_key IS NOT NULL AND source_key != ''"
        )
        self.conn.commit()

    def _add_column(self, table, column, declaration):
        if table not in TABLES or not SAFE_KEY.fullmatch(column):
            raise ValueError("Geçersiz alan.")
        names = {row[1] for row in self.conn.execute(f"PRAGMA table_info({table})")}
        if column not in names:
            self.conn.execute(f"ALTER TABLE {table} ADD COLUMN {column} {declaration}")

    def _table(self, table):
        if table not in TABLES:
            raise ValueError("Geçersiz tablo.")
        return table

    def _keys(self, data):
        keys = list(data)
        if not keys or any(not SAFE_KEY.fullmatch(key) for key in keys):
            raise ValueError("Geçersiz alan.")
        return keys

    def all(self, table):
        table = self._table(table)
        return self.conn.execute(f"SELECT * FROM {table} ORDER BY id DESC").fetchall()

    def get(self, table, row_id):
        table = self._table(table)
        return self.conn.execute(f"SELECT * FROM {table} WHERE id=?", (int(row_id),)).fetchone()

    def find_source(self, source_key):
        return self.conn.execute("SELECT id FROM invoices WHERE source_key=?", (source_key,)).fetchone()

    def save(self, table, data, row_id=None):
        table = self._table(table)
        keys = self._keys(data)
        values = [data[key] for key in keys]
        if row_id:
            assignments = ",".join(f"{key}=?" for key in keys)
            self.conn.execute(f"UPDATE {table} SET {assignments} WHERE id=?", values + [int(row_id)])
            saved_id = int(row_id)
        else:
            columns = ",".join(keys)
            marks = ",".join("?" for _ in keys)
            cursor = self.conn.execute(f"INSERT INTO {table} ({columns}) VALUES ({marks})", values)
            saved_id = int(cursor.lastrowid)
        self.conn.commit()
        return saved_id

    def delete(self, table, row_id):
        table = self._table(table)
        self.conn.execute(f"DELETE FROM {table} WHERE id=?", (int(row_id),))
        self.conn.commit()

    def offers_for(self, project_id):
        return self.conn.execute(
            "SELECT * FROM offers WHERE project_id=? ORDER BY amount IS NULL, amount, id DESC",
            (int(project_id),),
        ).fetchall()

    def next_issue_no(self, project_id):
        row = self.conn.execute(
            "SELECT COALESCE(MAX(issue_no), 0) + 1 AS n FROM tasks WHERE project_id=?",
            (int(project_id),),
        ).fetchone()
        return int(row["n"]) if row else 1

    def ensure_workspace(self):
        if getattr(self, "_workspace_ready", False):
            return
        taken = set()
        for project in sorted(self.all("projects"), key=lambda item: item["id"]):
            current = "".join(ch for ch in str(project["proj_key"] or "").upper() if ch.isalnum())[:6]
            if not current or current in taken:
                current = project_key(project["name"], taken)
                self.save("projects", {"proj_key": current}, project["id"])
            taken.add(current)
        projects = list(self.all("projects"))
        genel = next((project for project in projects if (project["proj_key"] or "").upper() == "GEN"), None)
        homeless = self.conn.execute(
            "SELECT 1 FROM tasks WHERE project_id IS NULL OR project_id = 0 OR project_id NOT IN (SELECT id FROM projects) LIMIT 1"
        ).fetchone()
        if genel is None and (homeless or not projects):
            genel_id = self.save("projects", {
                "name": "Genel",
                "proj_key": "GEN",
                "status": "Devam ediyor",
                "description": "Bir projeye bağlanmamış kayıtlar.",
            })
        elif genel is not None:
            genel_id = genel["id"]
        else:
            genel_id = projects[0]["id"]
        projects = {project["id"]: project for project in self.all("projects")}
        counters = {}
        for project_id in projects:
            row = self.conn.execute(
                "SELECT COALESCE(MAX(issue_no), 0) AS n FROM tasks WHERE project_id=?",
                (project_id,),
            ).fetchone()
            counters[project_id] = int(row["n"] or 0)
        orphans = self.conn.execute(
            """
            SELECT * FROM tasks
            WHERE project_id IS NULL OR project_id = 0 OR project_id NOT IN (SELECT id FROM projects)
               OR issue_no IS NULL OR issue_no = 0
               OR issue_type IS NULL OR issue_type = ''
            ORDER BY id
            """
        ).fetchall()
        for task in orphans:
            project_id = task["project_id"] if task["project_id"] in projects else genel_id
            updates = {}
            if task["project_id"] != project_id:
                updates["project_id"] = project_id
            if not task["issue_no"]:
                counters[project_id] = counters.get(project_id, 0) + 1
                updates["issue_no"] = counters[project_id]
            if not (task["issue_type"] or "").strip():
                updates["issue_type"] = "Görev"
            if updates:
                self.save("tasks", updates, task["id"])
        self._workspace_ready = True

    def offer_count(self, project_id):
        row = self.conn.execute("SELECT COUNT(*) AS n FROM offers WHERE project_id=?", (int(project_id),)).fetchone()
        return int(row["n"]) if row else 0

    def delete_project_offers(self, project_id):
        for row in self.offers_for(project_id):
            remove_offer_file(row["file_path"])
        self.conn.execute("DELETE FROM offers WHERE project_id=?", (int(project_id),))
        self.conn.commit()

    def ignore(self, key):
        self.conn.execute("INSERT OR IGNORE INTO ignored_sources(key) VALUES(?)", (key,))
        self.conn.commit()

    def ignored(self, key):
        return self.conn.execute("SELECT 1 FROM ignored_sources WHERE key=?", (key,)).fetchone() is not None

    def get_setting(self, key, default=""):
        if not SETTING_KEY.fullmatch(key):
            raise ValueError("Geçersiz ayar.")
        row = self.conn.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
        return default if row is None or row["value"] is None else row["value"]

    def set_setting(self, key, value):
        if not SETTING_KEY.fullmatch(key):
            raise ValueError("Geçersiz ayar.")
        self.conn.execute(
            "INSERT INTO settings(key, value) VALUES(?, ?) ON CONFLICT(key) DO UPDATE SET value=excluded.value",
            (key, value),
        )
        self.conn.commit()

    def seen(self, key):
        return self.conn.execute("SELECT 1 FROM notify_log WHERE key=?", (key,)).fetchone() is not None

    def mark_seen(self, key):
        self.conn.execute(
            "INSERT INTO notify_log(key, last_on) VALUES(?, ?) ON CONFLICT(key) DO UPDATE SET last_on=excluded.last_on",
            (key, today()),
        )
        self.conn.commit()

    def backup_to(self, path):
        dest = sqlite3.connect(path)
        try:
            self.conn.backup(dest)
        finally:
            dest.close()

    def close(self):
        self.conn.close()

def _round_rect(canvas, x1, y1, x2, y2, radius, **kwargs):
    radius = min(radius, (x2 - x1) / 2, (y2 - y1) / 2)
    style = kwargs.pop("style", "pieslice")
    canvas.create_arc(x1, y1, x1 + 2 * radius, y1 + 2 * radius, start=90, extent=90, style=style, **kwargs)
    canvas.create_arc(x2 - 2 * radius, y1, x2, y1 + 2 * radius, start=0, extent=90, style=style, **kwargs)
    canvas.create_arc(x1, y2 - 2 * radius, x1 + 2 * radius, y2, start=180, extent=90, style=style, **kwargs)
    canvas.create_arc(x2 - 2 * radius, y2 - 2 * radius, x2, y2, start=270, extent=90, style=style, **kwargs)
    canvas.create_rectangle(x1 + radius, y1, x2 - radius, y2, **kwargs)
    canvas.create_rectangle(x1, y1 + radius, x2, y2 - radius, **kwargs)


class Click(tk.Frame):
    """macOS Tk düğmesi renk vermez; bu yüzden tıklanan etiket kullanılır."""

    def __init__(self, master, text, command, bg, fg, hover=None, font=None, padx=14, pady=8, anchor="center", border=None):
        super().__init__(master, bg=bg, cursor="hand2", highlightthickness=1 if border else 0, highlightbackground=border or bg)
        self.command = command
        self._bg = bg
        self._fg = fg
        self._hover = hover or bg
        self._border = border
        self.label = tk.Label(self, text=text, bg=bg, fg=fg, font=font or (FONT, 13), padx=padx, pady=pady, anchor=anchor, cursor="hand2")
        self.label.pack(fill="both", expand=True)
        for widget in (self, self.label):
            widget.bind("<Enter>", self._on_enter)
            widget.bind("<Leave>", self._on_leave)
            widget.bind("<Button-1>", self._on_click)

    def _paint(self, bg):
        self.configure(bg=bg, highlightbackground=self._border or bg)
        self.label.configure(bg=bg, fg=self._fg)

    def _on_enter(self, _event):
        self._paint(self._hover)

    def _on_leave(self, _event):
        self._paint(self._bg)

    def _on_click(self, _event):
        if self.command:
            self.command()

    def set_text(self, text):
        self.label.configure(text=text)

    def set_palette(self, bg, fg, hover, border=None):
        self._bg = bg
        self._fg = fg
        self._hover = hover
        self._border = border
        self.configure(highlightthickness=1 if border else 0)
        self._paint(bg)


class NavItem(tk.Frame):
    def __init__(self, master, text, command):
        super().__init__(master, bg=NAV, cursor="hand2")
        self.click = Click(self, text, command, NAV, NAV_TEXT, NAV_HOVER, font=(FONT, 13), padx=14, pady=9, anchor="w")
        self.click.pack(fill="x", padx=10, pady=2)

    def set_active(self, active):
        if active:
            self.click.set_palette(NAV_ON, "white", NAV_ON)
        else:
            self.click.set_palette(NAV, NAV_TEXT, NAV_HOVER)


class CalendarPopup(tk.Toplevel):
    def __init__(self, parent, var):
        super().__init__(parent)
        self.var = var
        current = parse_day(var.get()) or date.today()
        self.year = current.year
        self.month = current.month
        self.selected = current
        self.title("Tarih seç")
        self.configure(bg=CANVAS)
        self.transient(parent)
        self.resizable(False, False)
        self.card = tk.Frame(self, bg=CARD, padx=14, pady=14, highlightbackground=LINE, highlightthickness=1)
        self.card.pack(padx=12, pady=12)
        self._draw()
        self.bind("<Escape>", lambda _e: self.destroy())
        self.geometry(f"+{parent.winfo_rootx() + 40}+{parent.winfo_rooty() + 40}")
        self.grab_set()

    def _draw(self):
        for child in self.card.winfo_children():
            child.destroy()
        header = tk.Frame(self.card, bg=CARD)
        header.pack(fill="x", pady=(0, 8))
        Click(header, "‹", self._prev, CARD, INK, HOVER, font=(FONT, 16, "bold"), padx=8, pady=2, border=LINE).pack(side="left")
        tk.Label(header, text=f"{MONTHS[self.month - 1]} {self.year}", bg=CARD, fg=INK, font=(FONT, 14, "bold")).pack(side="left", expand=True)
        Click(header, "›", self._next, CARD, INK, HOVER, font=(FONT, 16, "bold"), padx=8, pady=2, border=LINE).pack(side="right")
        grid = tk.Frame(self.card, bg=CARD)
        grid.pack()
        for col, name in enumerate("Pt Sa Ça Pe Cu Ct Pz".split()):
            tk.Label(grid, text=name, bg=CARD, fg=MUTED, font=(FONT, 11, "bold"), width=3).grid(row=0, column=col, pady=(0, 4))
        today_date = date.today()
        for row_index, week in enumerate(calendar.Calendar(0).monthdayscalendar(self.year, self.month), start=1):
            for col, day in enumerate(week):
                if not day:
                    tk.Label(grid, text="", bg=CARD, width=3).grid(row=row_index, column=col, padx=1, pady=1)
                    continue
                this = date(self.year, self.month, day)
                selected = this == self.selected
                current = this == today_date and not selected
                cell = tk.Label(
                    grid,
                    text=str(day),
                    bg=ACCENT if selected else (SOFT if current else CARD),
                    fg="white" if selected else INK,
                    font=(FONT, 12, "bold" if selected or current else "normal"),
                    width=3,
                    pady=4,
                    cursor="hand2",
                )
                cell.grid(row=row_index, column=col, padx=1, pady=1)
                cell.bind("<Button-1>", lambda _e, chosen=this: self._choose(chosen))

    def _shift(self, step):
        month = self.month + step
        year = self.year
        if month < 1:
            month, year = 12, year - 1
        elif month > 12:
            month, year = 1, year + 1
        self.month, self.year = month, year
        self._draw()

    def _prev(self):
        self._shift(-1)

    def _next(self):
        self._shift(1)

    def _choose(self, chosen):
        self.var.set(chosen.strftime(DATE_FMT))
        self.destroy()


class ClockPopup(tk.Toplevel):
    def __init__(self, parent, var):
        super().__init__(parent)
        self.var = var
        parsed = self._parse(var.get())
        now = datetime.now()
        self.hour = parsed[0] if parsed else now.hour
        self.minute = parsed[1] if parsed else now.minute
        self.mode = "hour"
        self.title("Saat seç")
        self.configure(bg=CANVAS)
        self.transient(parent)
        self.resizable(False, False)
        card = tk.Frame(self, bg=CARD, padx=16, pady=14, highlightbackground=LINE, highlightthickness=1)
        card.pack(padx=12, pady=12)
        self.caption = tk.Label(card, text="", bg=CARD, fg=INK, font=(FONT, 14, "bold"))
        self.caption.pack(anchor="w")
        self.canvas = tk.Canvas(card, width=240, height=240, bg=CARD, highlightthickness=0)
        self.canvas.pack(pady=(6, 8))
        self.canvas.bind("<Button-1>", self._click)
        periods = tk.Frame(card, bg=CARD)
        periods.pack()
        self.morning = Click(periods, "Öğleden önce", lambda: self._period(False), CARD, INK, HOVER, font=(FONT, 12), padx=10, pady=6, border=LINE)
        self.morning.pack(side="left", padx=(0, 6))
        self.evening = Click(periods, "Öğleden sonra", lambda: self._period(True), CARD, INK, HOVER, font=(FONT, 12), padx=10, pady=6, border=LINE)
        self.evening.pack(side="left")
        buttons = tk.Frame(card, bg=CARD)
        buttons.pack(anchor="e", pady=(12, 0))
        Click(buttons, "Vazgeç", self.destroy, CANVAS, INK, HOVER, font=(FONT, 12), padx=12, pady=6, border=LINE).pack(side="right")
        Click(buttons, "Tamam", self._ok, ACCENT, "white", ACCENT_HOVER, font=(FONT, 12, "bold"), padx=14, pady=6).pack(side="right", padx=(0, 8))
        self.bind("<Escape>", lambda _e: self.destroy())
        self._draw()
        self.geometry(f"+{parent.winfo_rootx() + 40}+{parent.winfo_rooty() + 20}")
        self.grab_set()

    def _parse(self, value):
        parts = str(value or "").strip().split(":")
        if len(parts) != 2 or not all(part.isdigit() for part in parts):
            return None
        hour, minute = int(parts[0]), int(parts[1])
        if hour > 23 or minute > 59:
            return None
        return hour, minute

    def _period(self, afternoon):
        self.hour = (self.hour % 12) + (12 if afternoon else 0)
        self._draw()

    def _click(self, event):
        dx, dy = event.x - 120, event.y - 120
        distance = math.hypot(dx, dy)
        if distance < 28 or distance > 112:
            return
        angle = math.atan2(dx, -dy)
        if angle < 0:
            angle += 2 * math.pi
        if self.mode == "hour":
            slot = int(angle / (2 * math.pi) * 12 + 0.5) % 12
            self.hour = slot + (12 if self.hour >= 12 else 0)
            self.mode = "minute"
        else:
            self.minute = int(angle / (2 * math.pi) * 60 + 0.5) % 60
        self._draw()

    def _hand(self, amount, length, color, width):
        angle = amount * 2 * math.pi - math.pi / 2
        x = 120 + math.cos(angle) * length
        y = 120 + math.sin(angle) * length
        self.canvas.create_line(120, 120, x, y, fill=color, width=width, capstyle="round")

    def _draw(self):
        self.caption.configure(text=f"{self.hour:02d}:{self.minute:02d}   ·   {'saati seç' if self.mode == 'hour' else 'dakikayı seç'}")
        afternoon = self.hour >= 12
        self.morning.set_palette(CARD if afternoon else ACCENT, INK if afternoon else "white", HOVER if afternoon else ACCENT_HOVER, LINE if afternoon else None)
        self.evening.set_palette(ACCENT if afternoon else CARD, "white" if afternoon else INK, ACCENT_HOVER if afternoon else HOVER, None if afternoon else LINE)
        self.canvas.delete("all")
        self.canvas.create_oval(18, 18, 222, 222, outline=LINE, width=2)
        for index in range(12):
            number = 12 if index == 0 else index
            angle = index / 12 * 2 * math.pi - math.pi / 2
            x = 120 + math.cos(angle) * 86
            y = 120 + math.sin(angle) * 86
            active = self.mode == "hour" and (self.hour % 12) == (0 if number == 12 else number)
            self.canvas.create_text(x, y, text=str(number), fill="white" if active else INK, font=(FONT, 13, "bold"))
            if active:
                self.canvas.create_oval(x - 14, y - 14, x + 14, y + 14, fill=ACCENT, outline=ACCENT)
                self.canvas.create_text(x, y, text=str(number), fill="white", font=(FONT, 13, "bold"))
        for index in range(60):
            if index % 5:
                continue
            angle = index / 60 * 2 * math.pi - math.pi / 2
            x = 120 + math.cos(angle) * 68
            y = 120 + math.sin(angle) * 68
            active = self.mode == "minute" and self.minute == index
            self.canvas.create_oval(x - 3, y - 3, x + 3, y + 3, fill=ACCENT if active else MUTED, outline="")
        self._hand((self.hour % 12) / 12 + self.minute / 60 / 12, 52, INK, 4)
        self._hand(self.minute / 60, 74, ACCENT, 2)
        self.canvas.create_oval(114, 114, 126, 126, fill=INK, outline=INK)

    def _ok(self):
        self.var.set(f"{self.hour:02d}:{self.minute:02d}")
        self.destroy()


class RecordDialog(tk.Toplevel):
    def __init__(self, parent, config, initial=None, watch_root=None):
        super().__init__(parent)
        self.result = None
        self.config_data = config
        self.watch_root = watch_root
        self.ledger_docs = []
        self._ledger_added = []
        self._ledger_removed = []
        self._drop_zone = None
        self._closing = False
        self.title(config["title"] + (" düzenle" if initial else " ekle"))
        self.configure(bg=CANVAS)
        self.transient(parent)
        self.resizable(False, False)
        self.vars = {}
        self.widgets = {}
        body = tk.Frame(self, bg=CARD, padx=18, pady=16, highlightbackground=LINE, highlightthickness=1)
        body.pack(fill="both", expand=True, padx=12, pady=12)
        body.columnconfigure(1, weight=1)
        offset = 0
        if config["title"] == "Fatura":
            offset = 1
            pay = tk.Frame(body, bg=CARD)
            pay.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 10))
            tk.Label(pay, text="Ödeme durumu", bg=CARD, fg=INK, font=(FONT, 13, "bold")).pack(anchor="w", pady=(0, 6))
            row = tk.Frame(pay, bg=CARD)
            row.pack(anchor="w")
            self.pay_choices = {}
            for label, bg, fg in (("Ödenmedi", "#fde2e1", DANGER), ("Ödendi", "#d9f2e6", OK)):
                chip = Click(row, label, lambda value=label: self.choose_payment(value), bg, fg, bg, font=(FONT, 13, "bold"), padx=16, pady=8)
                chip.pack(side="left", padx=(0, 8))
                self.pay_choices[label] = chip
        for index, field in enumerate(config["fields"]):
            index += offset
            label, key, kind, *extra = field
            tk.Label(body, text=label, bg=CARD, fg=MUTED, font=(FONT, 12)).grid(row=index, column=0, sticky="nw", padx=(0, 12), pady=6)
            value = ""
            if initial is not None and initial[key] is not None:
                value = initial[key]
                if kind == "number":
                    value = format_amount(value)
            if kind == "choice":
                choices = extra[0]
                var = tk.StringVar(value=str(value or choices[0]))
                widget = ttk.Combobox(body, textvariable=var, values=choices, state="readonly", width=36)
                widget.grid(row=index, column=1, sticky="ew", pady=6)
                self.vars[key] = var
            elif kind == "multiline":
                widget = tk.Text(body, width=38, height=3, wrap="word", font=(FONT, 13), relief="flat", highlightthickness=1, highlightbackground=LINE)
                widget.insert("1.0", str(value))
                widget.grid(row=index, column=1, sticky="ew", pady=6)
                self.widgets[key] = widget
            elif kind == "date":
                frame = tk.Frame(body, bg=CARD)
                frame.grid(row=index, column=1, sticky="ew", pady=6)
                var = tk.StringVar(value=str(value))
                tk.Entry(frame, textvariable=var, relief="flat", highlightthickness=1, highlightbackground=LINE, font=(FONT, 13)).pack(side="left", fill="x", expand=True, ipady=4)
                Click(frame, "Takvim", lambda current=var: self.open_picker(current, "date"), SOFT, INK, SOFT_HOVER, font=(FONT, 12), padx=10, pady=4).pack(side="left", padx=(6, 0))
                self.vars[key] = var
            elif kind == "time":
                var = tk.StringVar(value=str(value))
                widget = tk.Entry(body, textvariable=var, relief="flat", highlightthickness=1, highlightbackground=LINE, font=(FONT, 13))
                widget.grid(row=index, column=1, sticky="ew", pady=6, ipady=4)
                widget.bind("<Return>", lambda _e: self.on_ok())
                self.vars[key] = var
            elif kind == "category":
                frame = tk.Frame(body, bg=CARD)
                frame.grid(row=index, column=1, sticky="ew", pady=6)
                var = tk.StringVar(value=str(value))
                box = ttk.Combobox(frame, textvariable=var, state="readonly", width=28)
                box.pack(side="left", fill="x", expand=True)
                Click(frame, "+", self._add_ledger_category, ACCENT, "white", ACCENT_HOVER, font=(FONT, 14, "bold"), padx=10, pady=2).pack(side="left", padx=(6, 0))
                Click(frame, "−", self._remove_ledger_category, CARD, DANGER, "#fdecea", font=(FONT, 14, "bold"), padx=10, pady=2, border=LINE).pack(side="left", padx=(6, 0))
                self.vars[key] = var
                self.category_box = box
            elif kind == "file":
                frame = tk.Frame(body, bg=CARD)
                frame.grid(row=index, column=1, sticky="ew", pady=6)
                var = tk.StringVar(value=str(value))
                tk.Entry(frame, textvariable=var, relief="flat", highlightthickness=1, highlightbackground=LINE, font=(FONT, 12)).pack(side="left", fill="x", expand=True)
                Click(frame, "Dosya seç", lambda current=var: self.pick_file(current), SOFT, INK, SOFT_HOVER, font=(FONT, 12), padx=10, pady=4).pack(side="left", padx=(6, 0))
                self.vars[key] = var
            else:
                var = tk.StringVar(value=str(value))
                widget = tk.Entry(body, textvariable=var, relief="flat", highlightthickness=1, highlightbackground=LINE, font=(FONT, 13))
                widget.grid(row=index, column=1, sticky="ew", pady=6, ipady=4)
                widget.bind("<Return>", lambda _e: self.on_ok())
                self.vars[key] = var
        if config["title"] == "Fatura":
            self.mark_payment(self.vars["status"].get())
        if getattr(self, "category_box", None) is not None and "kind" in self.vars:
            self._sync_categories()
            self.vars["kind"].trace_add("write", lambda *_a: self._sync_categories())
        extra = 0
        if config["title"] == "Gelir / gider kaydı":
            extra = 1
            self._build_ledger_docs(body, len(config["fields"]) + offset, initial)
        buttons = tk.Frame(body, bg=CARD)
        buttons.grid(row=len(config["fields"]) + offset + extra, column=0, columnspan=2, sticky="e", pady=(12, 0))
        Click(buttons, "Vazgeç", self.destroy, CANVAS, INK, HOVER, font=(FONT, 12), padx=12, pady=6, border=LINE).pack(side="right")
        Click(buttons, "Kaydet", self.on_ok, ACCENT, "white", ACCENT_HOVER, font=(FONT, 12, "bold"), padx=14, pady=6).pack(side="right", padx=(0, 8))
        self.bind("<Escape>", lambda _e: self.destroy())
        self.update_idletasks()
        self.geometry(f"+{parent.winfo_rootx() + 90}+{parent.winfo_rooty() + 70}")
        self.grab_set()

    def choose_payment(self, value):
        self.vars["status"].set(value)
        self.mark_payment(value)
        self.on_ok()

    def mark_payment(self, value):
        for label, chip in getattr(self, "pay_choices", {}).items():
            if label == value:
                chip.set_palette(DANGER if label == "Ödenmedi" else ACCENT, "white", DANGER if label == "Ödenmedi" else ACCENT_HOVER)
            else:
                bg, fg = (DANGER_SOFT, DANGER) if label == "Ödenmedi" else (OK_SOFT, OK)
                chip.set_palette(bg, fg, bg)

    def _sync_categories(self):
        if getattr(self, "category_box", None) is None:
            return
        names = ledger_category_names(self.master.db, self.vars["kind"].get())
        current = self.vars["category"].get().strip()
        if current and fold(current) not in {fold(name) for name in names}:
            names = [current, *names]
        self.category_box["values"] = names

    def _add_ledger_category(self):
        name = simpledialog.askstring("Kategori", "Yeni kategori adı", parent=self)
        text = " ".join(str(name or "").split())[:40]
        if not text:
            return
        kind = self.vars["kind"].get()
        names = ledger_category_names(self.master.db, kind)
        for existing in names:
            if fold(existing) == fold(text):
                self.vars["category"].set(existing)
                self._sync_categories()
                return
        remember_ledger_category(self.master.db, kind, text)
        self.vars["category"].set(text)
        self._sync_categories()

    def _remove_ledger_category(self):
        kind = self.vars["kind"].get()
        current = self.vars["category"].get().strip()
        if not current:
            messagebox.showinfo(APP_NAME, "Silmek için önce bir kategori seç.", parent=self)
            return
        names = ledger_category_names(self.master.db, kind)
        match = next((name for name in names if fold(name) == fold(current)), "")
        if not match:
            return
        if not messagebox.askyesno(APP_NAME, f"“{match}” kategorisi listeden silinsin mi?", parent=self):
            return
        forget_ledger_category(self.master.db, kind, match)
        remaining = ledger_category_names(self.master.db, kind)
        self.vars["category"].set(remaining[0] if remaining else "")
        self._sync_categories()

    def open_picker(self, var, mode):
        if mode == "date":
            CalendarPopup(self, var)
        else:
            ClockPopup(self, var)

    def pick_file(self, var):
        path = filedialog.askopenfilename(
            parent=self,
            title="PDF veya görüntü seç",
            filetypes=[("Fatura", "*.pdf *.png *.jpg *.jpeg *.webp *.tif *.tiff *.heic"), ("Tümü", "*.*")],
        )
        if path:
            var.set(path)

    def _build_ledger_docs(self, body, row, initial):
        raw = ""
        if initial is not None and "docs_json" in initial.keys():
            raw = initial["docs_json"] or ""
        self.ledger_docs = []
        for item in ledger_documents(raw):
            path = Path(item["path"])
            try:
                resolved = path.resolve()
            except OSError:
                continue
            if path.is_symlink() or not resolved.is_file() or not is_within(resolved, LEDGER_DIR):
                continue
            self.ledger_docs.append({"path": str(resolved), "name": item["name"]})
        zone = tk.Frame(body, bg="#F4F8FF", highlightbackground="#1D6FEA", highlightthickness=1, padx=10, pady=8)
        zone.grid(row=row, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        self._drop_zone = zone
        tk.Label(zone, text="Faturalar", bg="#F4F8FF", fg=INK, font=(FONT, 13, "bold"), anchor="w").pack(fill="x")
        hint = "PDF veya görüntüyü bu alana bırak. Dosya ekle ile de seçebilirsin." if sys.platform == "darwin" else "PDF veya görüntü seç. Eklenen belge kayıttan çıkarılabilir."
        tk.Label(zone, text=hint, bg="#F4F8FF", fg=MUTED, font=(FONT, 11), anchor="w", justify="left", wraplength=420).pack(fill="x", pady=(2, 6))
        self._doc_rows = tk.Frame(zone, bg="#F4F8FF")
        self._doc_rows.pack(fill="x")
        Click(zone, "Dosya ekle", self._pick_ledger_files, ACCENT, "white", ACCENT_HOVER, font=(FONT, 12, "bold"), padx=12, pady=6).pack(anchor="w", pady=(6, 0))
        self._render_ledger_docs()
        self._drop_view = None
        self._drop_hot = False
        self.after(80, self._arm_file_drop)

    def _arm_file_drop(self):
        if self._closing:
            return
        self._drop_view = enable_mac_file_drop(self)
        self._poll_file_drop()

    def _poll_file_drop(self):
        if self._closing:
            return
        lib = _DROP_LIB
        if lib is not None:
            hot = bool(lib.ledger_hover())
            if hot != self._drop_hot:
                self._drop_hot = hot
                if hot:
                    self.entered()
                else:
                    self.exited()
            found = ctypes.create_string_buffer(48000)
            if lib.ledger_take(found, len(found)):
                paths = [line for line in found.value.decode("utf-8", "ignore").split("\n") if line]
                if paths:
                    self.dropped(paths)
        self.after(200, self._poll_file_drop)

    def entered(self):
        if self._drop_zone is not None:
            self._drop_zone.configure(highlightbackground="#E10600", highlightthickness=2)

    def exited(self):
        if self._drop_zone is not None:
            self._drop_zone.configure(highlightbackground="#1D6FEA", highlightthickness=1)

    def dropped(self, paths):
        self.exited()
        self._add_ledger_paths(paths)

    def _pick_ledger_files(self):
        paths = filedialog.askopenfilenames(
            parent=self,
            title="Fatura seç",
            filetypes=[("Fatura", "*.pdf *.png *.jpg *.jpeg *.webp *.tif *.tiff *.heic"), ("Tümü", "*.*")],
        )
        self._add_ledger_paths(paths)

    def _add_ledger_paths(self, paths):
        errors = []
        for path in paths or []:
            if len(self.ledger_docs) >= 30:
                errors.append("En fazla 30 belge eklenebilir.")
                break
            try:
                stored = store_ledger_file(path)
            except ValueError as exc:
                errors.append(f"{Path(path).name}: {exc}")
                continue
            if any(item["path"] == stored for item in self.ledger_docs):
                continue
            self.ledger_docs.append({"path": stored, "name": Path(path).name[:120]})
            if stored in self._ledger_removed:
                self._ledger_removed = [item for item in self._ledger_removed if item != stored]
            else:
                self._ledger_added.append(stored)
        self._render_ledger_docs()
        if errors:
            messagebox.showwarning(APP_NAME, "\n".join(errors[:6]), parent=self)

    def _render_ledger_docs(self):
        for child in self._doc_rows.winfo_children():
            child.destroy()
        if not self.ledger_docs:
            tk.Label(self._doc_rows, text="Henüz fatura yok.", bg="#F4F8FF", fg=MUTED, font=(FONT, 12), anchor="w").pack(fill="x")
            return
        for item in self.ledger_docs:
            line = tk.Frame(self._doc_rows, bg="#F4F8FF")
            line.pack(fill="x", pady=2)
            name = tk.Label(line, text=item["name"], bg="#F4F8FF", fg="#1D6FEA", font=(FONT, 12), anchor="w", cursor="hand2")
            name.pack(side="left", fill="x", expand=True)
            name.bind("<Button-1>", lambda _e, current=item: self._open_ledger_file(current["path"]))
            Click(line, "Çıkar", lambda current=item["path"]: self._remove_ledger_file(current), "#F4F8FF", DANGER, "#fdecea", font=(FONT, 11), padx=8, pady=2, border=LINE).pack(side="right")

    def _open_ledger_file(self, path):
        try:
            resolved = assert_openable(path, [LEDGER_DIR])
        except ValueError as exc:
            messagebox.showwarning(APP_NAME, str(exc), parent=self)
            return
        open_path(resolved)

    def _remove_ledger_file(self, path):
        self.ledger_docs = [item for item in self.ledger_docs if item["path"] != path]
        if path in self._ledger_added:
            remove_ledger_file(path)
            self._ledger_added = [item for item in self._ledger_added if item != path]
        elif path not in self._ledger_removed:
            self._ledger_removed.append(path)
        self._render_ledger_docs()

    def _discard_unsaved_docs(self):
        if self.result is not None:
            return
        for path in self._ledger_added:
            remove_ledger_file(path)
        self._ledger_added = []

    def destroy(self):
        if self._closing:
            return
        self._closing = True
        self._discard_unsaved_docs()
        clear_mac_file_drop(self)
        super().destroy()

    def on_ok(self):
        data = {}
        for field in self.config_data["fields"]:
            label, key, kind, *extra = field
            required = bool(extra) and extra[0] is True and kind != "choice"
            if kind == "multiline":
                value = self.widgets[key].get("1.0", "end").strip()
            else:
                value = self.vars[key].get().strip()
            if required and not value:
                messagebox.showwarning(APP_NAME, f"{label} boş bırakılamaz.", parent=self)
                return
            if kind == "date" and value and parse_day(value) is None:
                messagebox.showwarning(APP_NAME, f"{label} için YYYY-AA-GG kullan.", parent=self)
                return
            if kind == "time" and value:
                parts = value.split(":")
                if len(parts) != 2 or not all(part.isdigit() for part in parts):
                    messagebox.showwarning(APP_NAME, "Saat HH:MM biçiminde olmalı.", parent=self)
                    return
                hour, minute = int(parts[0]), int(parts[1])
                if hour > 23 or minute > 59:
                    messagebox.showwarning(APP_NAME, "Saat HH:MM biçiminde olmalı.", parent=self)
                    return
                value = f"{hour:02d}:{minute:02d}"
            if kind == "number":
                try:
                    value = parse_amount(value, required)
                except ValueError as exc:
                    messagebox.showwarning(APP_NAME, str(exc), parent=self)
                    return
            if kind == "file" and value:
                try:
                    value = store_attachment(value, self.watch_root)
                except ValueError as exc:
                    messagebox.showwarning(APP_NAME, str(exc), parent=self)
                    return
            data[key] = value
        if self.config_data["title"] == "Gelir / gider kaydı":
            data["docs_json"] = json.dumps(self.ledger_docs, ensure_ascii=False)
            for path in self._ledger_removed:
                remove_ledger_file(path)
            self._ledger_removed = []
            self._ledger_added = []
        self.result = data
        self.destroy()


def contact_fields(kind, scope):
    if kind == "Kişi" and scope == "Şirket içi":
        return [
            ("Ad soyad", "name", "text", True),
            ("Dahili numara", "extension", "text", False),
            ("Telefon", "phone", "text", False),
            ("E-posta", "email", "text", False),
            ("Notlar", "notes", "multiline", False),
        ]
    if kind == "Kişi":
        return [
            ("Ad soyad", "name", "text", True),
            ("Firma", "contact_person", "text", False),
            ("Telefon", "phone", "text", False),
            ("E-posta", "email", "text", False),
            ("Vergi / kimlik no", "tax_no", "text", False),
            ("Adres", "address", "multiline", False),
            ("Notlar", "notes", "multiline", False),
        ]
    return [
        ("Firma adı", "name", "text", True),
        ("Yetkili kişi", "contact_person", "text", False),
        ("Telefon", "phone", "text", False),
        ("E-posta", "email", "text", False),
        ("Vergi / kimlik no", "tax_no", "text", False),
        ("Adres", "address", "multiline", False),
        ("Notlar", "notes", "multiline", False),
    ]


class ContactDialog(tk.Toplevel):
    def __init__(self, parent, initial=None):
        super().__init__(parent)
        self.result = None
        self.initial = initial
        self.record_kind = self._cell("record_kind")
        self.scope = self._cell("scope")
        self.step = 0
        self.draft = {
            key: self._cell(key)
            for key in ("name", "contact_person", "phone", "email", "tax_no", "address", "notes", "extension")
        }
        self.vars = {}
        self.widgets = {}
        self.title("Kayıt türü")
        self.configure(bg=CANVAS)
        self.transient(parent)
        self.resizable(False, False)
        self.card = tk.Frame(self, bg=CARD, padx=18, pady=16, highlightbackground=LINE, highlightthickness=1)
        self.card.pack(fill="both", expand=True, padx=12, pady=12)
        self._show_step()
        self.bind("<Escape>", lambda _e: self.destroy())
        self.geometry(f"+{parent.winfo_rootx() + 120}+{parent.winfo_rooty() + 80}")
        self.grab_set()

    def _cell(self, key):
        row = self.initial
        if row is None or key not in row.keys() or row[key] is None:
            return ""
        return str(row[key])

    def _clear(self):
        for child in self.card.winfo_children():
            child.destroy()
        self.vars = {}
        self.widgets = {}

    def _show_step(self):
        self._clear()
        if self.step == 0:
            self.title("Kayıt türü")
            self._ask("Bu kayıt nedir?", "Önce kişi mi, firma mı olduğunu seç.", (("Kişi", "Kişi"), ("Firma", "Firma")), "record_kind")
        elif self.step == 1:
            self.title("Şirket içi mi?")
            self._ask("Şirket içi mi?", "Şirket içi kişilerde dahili numara sorulur. Şirket dışında mevcut alanlar durur.", (("Şirket içi", "Şirket içi"), ("Şirket dışı", "Şirket dışı")), "scope")
        else:
            self.title(f"{self.record_kind} · {self.scope}")
            self._form()

    def _ask(self, title, note, options, key):
        tk.Label(self.card, text=title, bg=CARD, fg=INK, font=(FONT, 16, "bold")).pack(anchor="w")
        tk.Label(self.card, text=note, bg=CARD, fg=MUTED, font=(FONT, 12), wraplength=420, justify="left").pack(anchor="w", pady=(4, 12))
        row = tk.Frame(self.card, bg=CARD)
        row.pack(anchor="w")
        current = getattr(self, key)
        for label, value in options:
            selected = value == current
            Click(
                row,
                label,
                lambda picked=value, field=key: self._pick(field, picked),
                ACCENT if selected else CARD,
                "white" if selected else INK,
                ACCENT_HOVER if selected else HOVER,
                font=(FONT, 13, "bold"),
                padx=18,
                pady=10,
                border=None if selected else LINE,
            ).pack(side="left", padx=(0, 8))
        buttons = tk.Frame(self.card, bg=CARD)
        buttons.pack(anchor="e", pady=(18, 0))
        if self.step:
            Click(buttons, "Geri", self._back, CANVAS, INK, HOVER, font=(FONT, 12), padx=12, pady=6, border=LINE).pack(side="right", padx=(8, 0))
        Click(buttons, "Devam", self._next, ACCENT, "white", ACCENT_HOVER, font=(FONT, 12, "bold"), padx=14, pady=6).pack(side="right")

    def _pick(self, key, value):
        setattr(self, key, value)
        self.after_idle(self._show_step)

    def _next(self):
        if self.step == 0 and not self.record_kind:
            messagebox.showwarning(APP_NAME, "Önce kişi veya firma seç.", parent=self)
            return
        if self.step == 1 and not self.scope:
            messagebox.showwarning(APP_NAME, "Şirket içi mi, şirket dışı mı seç.", parent=self)
            return
        self.step += 1
        self.after_idle(self._show_step)

    def _back(self):
        if self.step == 2:
            self._remember_form()
        self.step -= 1
        self.after_idle(self._show_step)

    def _form(self):
        fields = contact_fields(self.record_kind, self.scope)
        tk.Label(self.card, text=f"{self.record_kind} · {self.scope}", bg=CARD, fg=INK, font=(FONT, 16, "bold")).grid(row=0, column=0, columnspan=2, sticky="w")
        self.card.columnconfigure(1, weight=1)
        for index, (label, key, kind, required) in enumerate(fields, start=1):
            tk.Label(self.card, text=label, bg=CARD, fg=MUTED, font=(FONT, 12)).grid(row=index, column=0, sticky="nw", padx=(0, 12), pady=6)
            value = self.draft.get(key, "")
            if kind == "multiline":
                widget = tk.Text(self.card, width=38, height=3, wrap="word", font=(FONT, 13), relief="flat", highlightthickness=1, highlightbackground=LINE)
                widget.insert("1.0", value)
                widget.grid(row=index, column=1, sticky="ew", pady=6)
                self.widgets[key] = widget
            else:
                var = tk.StringVar(value=value)
                widget = tk.Entry(self.card, textvariable=var, relief="flat", highlightthickness=1, highlightbackground=LINE, font=(FONT, 13))
                widget.grid(row=index, column=1, sticky="ew", pady=6, ipady=4)
                widget.bind("<Return>", lambda _e: self._save())
                self.vars[key] = var
            if required:
                self.required_name = (label, key)
        buttons = tk.Frame(self.card, bg=CARD)
        buttons.grid(row=len(fields) + 1, column=0, columnspan=2, sticky="e", pady=(12, 0))
        Click(buttons, "Geri", self._back, CANVAS, INK, HOVER, font=(FONT, 12), padx=12, pady=6, border=LINE).pack(side="right")
        Click(buttons, "Kaydet", self._save, ACCENT, "white", ACCENT_HOVER, font=(FONT, 12, "bold"), padx=14, pady=6).pack(side="right", padx=(0, 8))

    def _remember_form(self):
        for key, widget in self.widgets.items():
            self.draft[key] = widget.get("1.0", "end").strip()
        for key, var in self.vars.items():
            self.draft[key] = var.get().strip()

    def _save(self):
        self._remember_form()
        fields = contact_fields(self.record_kind, self.scope)
        data = {key: "" for key in ("name", "contact_person", "phone", "email", "tax_no", "address", "notes", "extension")}
        for label, key, _kind, required in fields:
            value = self.draft.get(key, "").strip()
            if required and not value:
                messagebox.showwarning(APP_NAME, f"{label} boş bırakılamaz.", parent=self)
                return
            data[key] = value
        if not (self.record_kind == "Kişi" and self.scope == "Şirket içi"):
            data["extension"] = ""
        if self.record_kind == "Kişi" and self.scope == "Şirket içi":
            data["contact_person"] = ""
            data["tax_no"] = ""
            data["address"] = ""
        data["record_kind"] = self.record_kind
        data["scope"] = self.scope
        self.result = data
        self.destroy()


class ReportDialog(tk.Toplevel):
    def __init__(self, parent, report=None):
        super().__init__(parent)
        self.app = parent
        self.result = None
        self.report = report
        self.invoices = parent.db.all("invoices")
        self.row = self.invoices[0] if self.invoices else None
        self.files = []
        self.existing_pdf = ""
        if report is not None:
            chosen = next((item for item in self.invoices if item["id"] == report["invoice_id"]), None)
            if chosen is not None:
                self.row = chosen
                self.files = [line for line in str(chosen["report_files"] or "").splitlines() if line.strip()]
            self.existing_pdf = report["saved_path"] or ""
        self.title("Raporu düzenle" if report else "Rapor hazırla")
        self.configure(bg=CANVAS)
        self.transient(parent)
        self.resizable(False, False)
        card = tk.Frame(self, bg=CARD, padx=18, pady=16, highlightbackground=LINE, highlightthickness=1)
        card.pack(fill="both", expand=True, padx=12, pady=12)
        tk.Label(card, text="Fatura", bg=CARD, fg=INK, font=(FONT, 13, "bold")).pack(anchor="w")
        self.picker_open = False
        self.picker_box = tk.Frame(card, bg=CARD, highlightbackground=LINE, highlightthickness=1, cursor="hand2")
        self.picker_box.pack(fill="x", pady=(4, 10))
        self.picker_text = tk.Label(self.picker_box, text=self._invoice_label(self.row), bg=CARD, fg=INK, font=(FONT, 12), anchor="w", justify="left", wraplength=400, cursor="hand2")
        self.picker_text.pack(side="left", fill="x", expand=True, padx=10, pady=8)
        self.picker_arrow = tk.Label(self.picker_box, text="▾", bg=CARD, fg=MUTED, font=(FONT, 14), cursor="hand2", padx=8)
        self.picker_arrow.pack(side="right")
        for widget in (self.picker_box, self.picker_text, self.picker_arrow):
            widget.bind("<Button-1>", lambda _e: self._toggle_picker())
        self.picker_panel = tk.Frame(card, bg=CARD, highlightbackground=LINE, highlightthickness=1)
        search_row = tk.Frame(self.picker_panel, bg=CARD)
        search_row.pack(fill="x", padx=8, pady=(8, 4))
        tk.Label(search_row, text="Ara", bg=CARD, fg=MUTED, font=(FONT, 12)).pack(side="left", padx=(4, 6))
        self.invoice_query = tk.StringVar()
        self.invoice_query.trace_add("write", lambda *_: self._fill_invoice_list())
        self.invoice_search = tk.Entry(search_row, textvariable=self.invoice_query, relief="flat", highlightthickness=1, highlightbackground=LINE, font=(FONT, 12))
        self.invoice_search.pack(side="left", fill="x", expand=True, ipady=4)
        self.invoice_search.bind("<Escape>", lambda _e: self._close_picker())
        self.invoice_canvas = tk.Canvas(self.picker_panel, bg=CARD, height=168, highlightthickness=0)
        self.invoice_scroll = ttk.Scrollbar(self.picker_panel, orient="vertical", command=self.invoice_canvas.yview)
        self.invoice_canvas.configure(yscrollcommand=self.invoice_scroll.set)
        self.invoice_canvas.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=(0, 8))
        self.invoice_scroll.pack(side="right", fill="y", pady=(0, 8), padx=(0, 4))
        self.invoice_items = tk.Frame(self.invoice_canvas, bg=CARD)
        self.invoice_window = self.invoice_canvas.create_window((0, 0), window=self.invoice_items, anchor="nw")
        self.invoice_items.bind("<Configure>", lambda _e: self.invoice_canvas.configure(scrollregion=self.invoice_canvas.bbox("all")))
        self.invoice_canvas.bind("<Configure>", lambda e: self.invoice_canvas.itemconfigure(self.invoice_window, width=e.width))
        self.summary = tk.Label(card, text="", bg=CARD, fg=MUTED, font=(FONT, 12), wraplength=460, justify="left")
        self.summary.pack(anchor="w", pady=(0, 10))
        self._show_invoice()
        tk.Label(card, text="Harcama nedeni", bg=CARD, fg=INK, font=(FONT, 13, "bold")).pack(anchor="w")
        tk.Label(card, text="Satın alma ve ödemenin neden yapıldığını yaz.", bg=CARD, fg=MUTED, font=(FONT, 12)).pack(anchor="w", pady=(2, 6))
        self.reason = tk.Text(card, width=52, height=5, wrap="word", font=(FONT, 13), relief="flat", highlightthickness=1, highlightbackground=LINE)
        self.reason.pack(fill="x")
        if report is not None and report["reason"]:
            self.reason.insert("1.0", report["reason"])
        tk.Label(card, text="Ekler", bg=CARD, fg=INK, font=(FONT, 13, "bold")).pack(anchor="w", pady=(12, 0))
        tk.Label(card, text="Ekran görüntüsü, e-posta ve PDF tek dosyada birleşir.", bg=CARD, fg=MUTED, font=(FONT, 12)).pack(anchor="w", pady=(2, 6))
        adds = tk.Frame(card, bg=CARD)
        adds.pack(anchor="w")
        Click(adds, "Ekran görüntüsü", lambda: self._add("image"), SOFT, INK, SOFT_HOVER, font=(FONT, 12), padx=10, pady=6).pack(side="left")
        Click(adds, "E-posta", lambda: self._add("mail"), SOFT, INK, SOFT_HOVER, font=(FONT, 12), padx=10, pady=6).pack(side="left", padx=6)
        Click(adds, "PDF", lambda: self._add("pdf"), SOFT, INK, SOFT_HOVER, font=(FONT, 12), padx=10, pady=6).pack(side="left")
        self.file_box = tk.Frame(card, bg=CARD)
        self.file_box.pack(fill="x", pady=(8, 0))
        self._paint_files()
        buttons = tk.Frame(card, bg=CARD)
        buttons.pack(anchor="e", pady=(14, 0))
        Click(buttons, "Vazgeç", self.destroy, CANVAS, INK, HOVER, font=(FONT, 12), padx=12, pady=6, border=LINE).pack(side="right")
        Click(buttons, "Kaydet", self._save, ACCENT, "white", ACCENT_HOVER, font=(FONT, 12, "bold"), padx=14, pady=6).pack(side="right", padx=(0, 8))
        self.bind("<Escape>", self._on_escape)
        self.geometry(f"+{parent.winfo_rootx() + 80}+{parent.winfo_rooty() + 50}")
        self.grab_set()

    def _invoice_label(self, row):
        amount = format_amount(row["amount"])
        return f"{row['party'] or 'Fatura'}   ·   {row['invoice_no'] or 'no yok'}   ·   {amount + ' ₺' if amount else 'tutar yok'}"

    def _show_invoice(self):
        if self.row is None:
            self.picker_text.configure(text="Fatura yok")
            self.summary.configure(text="")
            return
        amount = format_amount(self.row["amount"])
        self.picker_text.configure(text=self._invoice_label(self.row))
        self.summary.configure(text=f"{amount + ' ₺' if amount else 'tutar yok'}   ·   {self.row['invoice_date'] or 'tarih yok'}   ·   {self.row['status'] or ''}")

    def _toggle_picker(self):
        if self.picker_open:
            self._close_picker()
            return "break"
        self.picker_panel.pack(fill="x", pady=(0, 8), before=self.summary)
        self.picker_arrow.configure(text="▴")
        self.picker_open = True
        self._fill_invoice_list()
        self.invoice_search.focus_set()
        return "break"

    def _close_picker(self):
        if not self.picker_open:
            return "break"
        self.picker_panel.pack_forget()
        self.picker_arrow.configure(text="▾")
        self.picker_open = False
        return "break"

    def _on_escape(self, _event=None):
        if self.picker_open:
            return self._close_picker()
        self.destroy()

    def _fill_invoice_list(self):
        for child in self.invoice_items.winfo_children():
            child.destroy()
        query = fold(self.invoice_query.get().strip())
        matches = []
        for row in self.invoices:
            amount = format_amount(row["amount"])
            blob = fold(" ".join((
                row["party"] or "",
                row["invoice_no"] or "",
                row["invoice_date"] or "",
                row["status"] or "",
                amount,
                amount.replace(".", "") if amount else "",
            )))
            if query and query not in blob:
                continue
            matches.append(row)
        if not matches:
            tk.Label(self.invoice_items, text="Sonuç yok", bg=CARD, fg=MUTED, font=(FONT, 12), anchor="w").pack(fill="x", padx=8, pady=8)
            return
        for row in matches:
            selected = row["id"] == self.row["id"]
            item = tk.Frame(self.invoice_items, bg=SOFT if selected else CARD, cursor="hand2")
            item.pack(fill="x")
            label = tk.Label(item, text=self._invoice_label(row), bg=item["bg"], fg=INK, font=(FONT, 12), anchor="w", justify="left", wraplength=400, cursor="hand2")
            label.pack(fill="x", padx=8, pady=6)
            for widget in (item, label):
                widget.bind("<Button-1>", lambda _e, chosen=row: self._pick_invoice(chosen))
                widget.bind("<Enter>", lambda _e, box=item, text=label: (box.configure(bg=HOVER), text.configure(bg=HOVER)))
                widget.bind("<Leave>", lambda _e, box=item, text=label, bg=SOFT if selected else CARD: (box.configure(bg=bg), text.configure(bg=bg)))

    def _pick_invoice(self, row):
        self.row = row
        self._show_invoice()
        self._close_picker()

    def _add(self, kind):
        chosen = {
            "image": ("Ekran görüntüsü seç", [("Görüntü", "*.png *.jpg *.jpeg *.webp *.heic *.tif *.tiff")]),
            "mail": ("E-posta dosyası seç", [("E-posta", "*.eml")]),
            "pdf": ("PDF seç", [("PDF", "*.pdf")]),
        }[kind]
        path = filedialog.askopenfilename(parent=self, title=chosen[0], filetypes=chosen[1])
        if not path:
            return
        try:
            stored = store_report_file(path)
        except ValueError as exc:
            messagebox.showwarning(APP_NAME, str(exc), parent=self)
            return
        if stored not in self.files:
            self.files.append(stored)
        self._paint_files()

    def _paint_files(self):
        for child in self.file_box.winfo_children():
            child.destroy()
        if not self.files:
            tk.Label(self.file_box, text="Henüz ek yok.", bg=CARD, fg=MUTED, font=(FONT, 12)).pack(anchor="w")
            return
        labels = {".eml": "E-posta", ".pdf": "PDF"}
        for path in self.files:
            ext = Path(path).suffix.lower()
            kind = labels.get(ext, "Görüntü")
            row = tk.Frame(self.file_box, bg="#f6f8fb", highlightbackground=LINE, highlightthickness=1)
            row.pack(fill="x", pady=3)
            tk.Label(row, text=f"{kind}   ·   {Path(path).name}", bg="#f6f8fb", fg=INK, font=(FONT, 12), anchor="w").pack(side="left", fill="x", expand=True, padx=8, pady=6)
            Click(row, "Kaldır", lambda item=path: self._remove(item), CARD, DANGER, "#fdecea", font=(FONT, 11), padx=8, pady=4, border=LINE).pack(side="right", padx=6, pady=4)

    def _remove(self, path):
        self.files = [item for item in self.files if item != path]
        self._paint_files()

    def _save(self):
        reason = self.reason.get("1.0", "end").strip()
        if len(reason) < 8:
            messagebox.showwarning(APP_NAME, "Satın alma ve ödemenin nedenini yaz.", parent=self)
            return
        if self.row is None:
            messagebox.showwarning(APP_NAME, "Rapor için bir fatura seç.", parent=self)
            return
        initial = safe_name(f"Rapor {self.row['party'] or 'fatura'} {self.row['invoice_date'] or today()}") + ".pdf"
        initialdir = ""
        if self.existing_pdf:
            current = Path(self.existing_pdf)
            if current.name:
                initial = current.name
            if current.parent.is_dir():
                initialdir = str(current.parent)
        path = filedialog.asksaveasfilename(parent=self, title="PDF nereye kaydedilsin?", defaultextension=".pdf", initialfile=initial, initialdir=initialdir, filetypes=[("PDF", "*.pdf")])
        if not path:
            return
        files = []
        attachment = self.row["attachment"] or ""
        if attachment:
            files.append(attachment)
        files.extend(item for item in self.files if item not in files)
        self.title("PDF hazırlanıyor…")
        self.update_idletasks()
        try:
            saved = write_report_pdf(path, self.row, reason, files)
        except (OSError, ValueError, subprocess.SubprocessError) as exc:
            self.title("Raporu düzenle" if self.report else "Rapor hazırla")
            messagebox.showwarning(APP_NAME, f"PDF birleştirilemedi.\n{exc}", parent=self)
            return
        self.result = {"report_reason": reason, "report_files": "\n".join(self.files), "saved_path": str(saved)}
        self.destroy()


class SettingsDialog(tk.Toplevel):
    def __init__(self, parent, values):
        super().__init__(parent)
        self.result = None
        self.title("E-posta ve klasör")
        self.configure(bg=CANVAS)
        self.transient(parent)
        self.resizable(False, False)
        self.vars = {key: tk.StringVar(value=str(value or "")) for key, value in values.items()}
        self.enabled = tk.BooleanVar(value=values.get("mail_enabled") == "1")
        card = tk.Frame(self, bg=CARD, padx=18, pady=16, highlightbackground=LINE, highlightthickness=1)
        card.pack(fill="both", expand=True, padx=12, pady=12)
        tk.Label(card, text="Ortak fatura klasörü", bg=CARD, fg=INK, font=(FONT, 15, "bold")).grid(row=0, column=0, columnspan=3, sticky="w")
        tk.Label(card, text="Bu klasöre bırakılan PDF ve görüntüler uygulamaya ve Excel listesine eklenir.", bg=CARD, fg=MUTED, font=(FONT, 12)).grid(row=1, column=0, columnspan=3, sticky="w", pady=(2, 8))
        self._path_row(card, 2, "Klasör", "watch_dir")
        self._path_row(card, 3, "Excel", "excel_path")
        tk.Label(card, text="E-posta", bg=CARD, fg=INK, font=(FONT, 15, "bold")).grid(row=4, column=0, columnspan=3, sticky="w", pady=(16, 2))
        tk.Label(card, text="Mac’teki Mail’i veya kullandığın programı seç. Son 21 gündeki fatura ekleri okunur.", bg=CARD, fg=MUTED, font=(FONT, 12), wraplength=520, justify="left").grid(row=5, column=0, columnspan=3, sticky="w", pady=(0, 8))
        programs = tk.Frame(card, bg=CARD)
        programs.grid(row=6, column=0, columnspan=3, sticky="w")
        self.provider_buttons = {}
        saved_host = self.vars["mail_host"].get().strip().lower()
        self.provider = "Mail" if sys.platform == "darwin" else "Gmail"
        for name, (host, _port, _box, _hint) in MAIL_PROGRAMS.items():
            if host and host.lower() == saved_host:
                self.provider = name
        for name in MAIL_PROGRAMS:
            if name == "Mail" and sys.platform != "darwin":
                continue
            button = Click(programs, name, lambda picked=name: self._choose_provider(picked), CARD, INK, HOVER, font=(FONT, 12), padx=10, pady=6, border=LINE)
            button.pack(side="left", padx=(0, 6), pady=(0, 6))
            self.provider_buttons[name] = button
        self.user_label = tk.Label(card, text="E-posta adresi", bg=CARD, fg=MUTED, font=(FONT, 12))
        self.user_label.grid(row=7, column=0, sticky="w", pady=4)
        self.user_entry = tk.Entry(card, textvariable=self.vars["mail_user"], width=36, relief="flat", highlightthickness=1, highlightbackground=LINE, font=(FONT, 13))
        self.user_entry.grid(row=7, column=1, columnspan=2, sticky="ew", pady=4, ipady=3)
        self.pass_label = tk.Label(card, text="Parola", bg=CARD, fg=MUTED, font=(FONT, 12))
        self.pass_label.grid(row=8, column=0, sticky="w", pady=4)
        self.password = tk.Entry(card, show="•", width=36, relief="flat", highlightthickness=1, highlightbackground=LINE, font=(FONT, 13))
        self.password.grid(row=8, column=1, columnspan=2, sticky="ew", pady=4, ipady=3)
        self.mail_hint = tk.Label(card, text="", bg=CARD, fg=MUTED, font=(FONT, 11), wraplength=460, justify="left")
        self.mail_hint.grid(row=9, column=0, columnspan=3, sticky="w", pady=(6, 0))
        self.advanced = tk.Frame(card, bg=CARD)
        self.advanced.grid(row=10, column=0, columnspan=3, sticky="ew")
        for offset, (label, key) in enumerate((("Sunucu", "mail_host"), ("Klasör", "mail_box"))):
            tk.Label(self.advanced, text=label, bg=CARD, fg=MUTED, font=(FONT, 12)).grid(row=offset, column=0, sticky="w", pady=4)
            tk.Entry(self.advanced, textvariable=self.vars[key], width=36, relief="flat", highlightthickness=1, highlightbackground=LINE, font=(FONT, 13)).grid(row=offset, column=1, sticky="ew", pady=4, ipady=3)
        self.advanced.columnconfigure(1, weight=1)
        buttons = tk.Frame(card, bg=CARD)
        buttons.grid(row=11, column=0, columnspan=3, sticky="e", pady=(14, 0))
        Click(buttons, "Vazgeç", self.destroy, CANVAS, INK, HOVER, font=(FONT, 12), padx=12, pady=6, border=LINE).pack(side="right")
        Click(buttons, "Kaydet", lambda: self._finish(False), CARD, INK, HOVER, font=(FONT, 12), padx=12, pady=6, border=LINE).pack(side="right", padx=(0, 8))
        Click(buttons, "E-postayı oku", lambda: self._finish(True), ACCENT, "white", ACCENT_HOVER, font=(FONT, 12, "bold"), padx=14, pady=6).pack(side="right", padx=(0, 8))
        self._choose_provider(self.provider)
        self.bind("<Escape>", lambda _e: self.destroy())
        self.geometry(f"+{parent.winfo_rootx() + 120}+{parent.winfo_rooty() + 60}")
        self.grab_set()

    def _path_row(self, parent, row, label, key):
        tk.Label(parent, text=label, bg=CARD, fg=MUTED, font=(FONT, 12)).grid(row=row, column=0, sticky="w", pady=4)
        tk.Entry(parent, textvariable=self.vars[key], width=42, relief="flat", highlightthickness=1, highlightbackground=LINE, font=(FONT, 12)).grid(row=row, column=1, sticky="ew", pady=4, ipady=3)
        chooser = filedialog.askdirectory if key == "watch_dir" else filedialog.askopenfilename
        Click(parent, "Seç", lambda: self._browse(key, chooser), SOFT, INK, SOFT_HOVER, font=(FONT, 12), padx=10, pady=4).grid(row=row, column=2, padx=(6, 0))

    def _browse(self, key, chooser):
        if key == "watch_dir":
            selected = chooser(parent=self, title="Fatura klasörü")
        else:
            selected = chooser(parent=self, title="Excel dosyası", filetypes=[("Excel", "*.xlsx")])
        if selected:
            self.vars[key].set(selected)

    def _show_credentials(self, visible):
        widgets = (self.user_label, self.user_entry, self.pass_label, self.password)
        if visible:
            for widget in widgets:
                widget.grid()
        else:
            for widget in widgets:
                widget.grid_remove()

    def _choose_provider(self, name):
        self.provider = name
        host, port, box, hint = MAIL_PROGRAMS[name]
        for label, button in self.provider_buttons.items():
            if label == name:
                button.set_palette(ACCENT, "white", ACCENT_HOVER)
            else:
                button.set_palette(CARD, INK, HOVER, LINE)
        if name == "Mail":
            self.vars["mail_host"].set(host)
            self.vars["mail_port"].set(port)
            self.vars["mail_box"].set(box)
            self._show_credentials(False)
            self.advanced.grid_remove()
            self.mail_hint.configure(text=hint)
            return
        self._show_credentials(True)
        if name != "Diğer":
            self.vars["mail_host"].set(host)
            self.vars["mail_port"].set(port)
            self.vars["mail_box"].set(box)
            self.advanced.grid_remove()
        else:
            self.vars["mail_port"].set(self.vars["mail_port"].get().strip() or port)
            self.vars["mail_box"].set(self.vars["mail_box"].get().strip() or box)
            self.advanced.grid()
        self.mail_hint.configure(text=hint + " Parola veritabanına yazılmaz. Boş bırakırsan kayıtlı parola durur.")

    def _collected(self):
        watch = Path(self.vars["watch_dir"].get().strip()).expanduser()
        excel = Path(self.vars["excel_path"].get().strip()).expanduser()
        if not acceptable_watch_dir(watch):
            raise ValueError("Klasör, ev dizininin içinde ayrı bir fatura klasörü olmalı. Masaüstünün tamamı seçilemez.")
        if excel.suffix.lower() != ".xlsx" or excel.is_symlink() or not is_within(excel if excel.exists() else excel.parent, watch):
            raise ValueError("Excel dosyası seçilen fatura klasörünün içinde .xlsx olmalı.")
        if self.provider == "Mail":
            if sys.platform != "darwin":
                raise ValueError("Mac’teki Mail yalnız Mac bilgisayarda kullanılır.")
            return {
                "watch_dir": str(watch),
                "excel_path": str(excel),
                "mail_enabled": "1" if self.enabled.get() else "0",
                "mail_host": LOCAL_MAIL_HOST,
                "mail_port": "0",
                "mail_user": "",
                "mail_box": "INBOX",
                "password": "",
            }
        host, port, box, _hint = MAIL_PROGRAMS[self.provider]
        if self.provider == "Diğer":
            host = self.vars["mail_host"].get().strip()
            port = self.vars["mail_port"].get().strip() or "993"
            box = self.vars["mail_box"].get().strip() or "INBOX"
        user = self.vars["mail_user"].get().strip()
        password = self.password.get()
        if self.enabled.get():
            if not port.isdigit() or not valid_host(host) or not valid_user(user) or not valid_mailbox(box):
                raise ValueError("E-posta adresi geçersiz. Kurumsal posta seçtiysen sunucuyu da yaz.")
            if password and not valid_password(password):
                raise ValueError("Parola boşluk satırı içeremez.")
            if not password and not keychain_get(user):
                raise ValueError("İlk okumada e-posta parolası gerekli.")
        return {
            "watch_dir": str(watch),
            "excel_path": str(excel),
            "mail_enabled": "1" if self.enabled.get() else "0",
            "mail_host": host,
            "mail_port": port,
            "mail_user": user,
            "mail_box": box,
            "password": password,
        }

    def _finish(self, read_now):
        if self.provider == "Mail" or read_now or self.vars["mail_user"].get().strip():
            self.enabled.set(True)
        else:
            self.enabled.set(False)
        try:
            self.result = self._collected()
        except ValueError as exc:
            messagebox.showwarning(APP_NAME, str(exc), parent=self)
            return
        self.result["read_now"] = read_now
        self.destroy()

    def on_test(self):
        try:
            data = self._collected()
        except ValueError as exc:
            messagebox.showwarning(APP_NAME, str(exc), parent=self)
            return
        password = data["password"] or keychain_get(data["mail_user"])
        try:
            client = imaplib.IMAP4_SSL(data["mail_host"], int(data["mail_port"]), timeout=15)
            try:
                client.login(data["mail_user"], password)
                status, _ = client.select(data["mail_box"], readonly=True)
                if status != "OK":
                    raise ValueError("Posta klasörü açılamadı.")
            finally:
                try:
                    client.logout()
                except imaplib.IMAP4.error:
                    pass
        except Exception:
            messagebox.showwarning(APP_NAME, "Posta okunamadı. Adresi ve parolayı kontrol et. Gmail ve iCloud uygulama parolası ister.", parent=self)
            return
        messagebox.showinfo(APP_NAME, "Bağlantı başarılı. Posta kutusu salt okunur açıldı.", parent=self)

    def on_ok(self):
        self._finish(False)


def fit_note_photo(path):
    image = tk.PhotoImage(file=str(path))
    scale = 1
    while (image.width() // scale > 360 or image.height() // scale > 280) and scale < 12:
        scale += 1
    if scale > 1:
        image = image.subsample(scale, scale)
    return image, image.width(), image.height()


class TaskNoteDialog(tk.Toplevel):
    def __init__(self, parent, title, raw, task_id):
        super().__init__(parent)
        self.task_title = title or "Görev"
        self.task_id = int(task_id)
        self.photos = []
        self.images = {}
        self.shapes = {}
        self.saved = None
        self.keep = False
        self.size = 18
        self.color = NOTE_COLORS[0]
        self.bold = False
        self.italic = False
        self.underline = False
        self.strike = False
        self.highlight = ""
        self.align = "left"
        self.indent = 0
        self.size_buttons = {}
        self.color_buttons = {}
        self.mark_buttons = {}
        self.align_buttons = {}
        self.hilite_buttons = {}
        self.title(f"Yorum · {self.task_title}")
        self.configure(bg=CANVAS)
        self.transient(parent)
        self.resizable(False, False)
        card = tk.Frame(self, bg=CARD, padx=16, pady=14, highlightbackground=LINE, highlightthickness=1)
        card.pack(fill="both", expand=True, padx=12, pady=12)
        tk.Label(card, text="Yorum", bg=CARD, fg=INK, font=(FONT, 16, "bold")).pack(anchor="w")
        tk.Label(card, text="Doğrudan yaz. Kalın, italik, liste ve hiza Word’deki gibi seçilir.", bg=CARD, fg=MUTED, font=(FONT, 12)).pack(anchor="w", pady=(2, 8))
        marks = tk.Frame(card, bg=CARD)
        marks.pack(fill="x")
        for key, label, font in (
            ("bold", "K", (FONT, 13, "bold")),
            ("italic", "İ", (FONT, 13, "italic")),
            ("underline", "A̲", (FONT, 13)),
            ("strike", "Ü", (FONT, 13)),
        ):
            button = Click(marks, label, lambda picked=key: self._toggle(picked), CARD, INK, HOVER, font=font, padx=10, pady=5, border=LINE)
            button.pack(side="left", padx=(0, 4))
            self.mark_buttons[key] = button
        tk.Label(marks, text="Vurgu", bg=CARD, fg=MUTED, font=(FONT, 12)).pack(side="left", padx=(8, 4))
        clear_mark = Click(marks, "Yok", lambda: self._pick_highlight(""), CARD, INK, HOVER, font=(FONT, 11), padx=8, pady=5, border=LINE)
        clear_mark.pack(side="left", padx=(0, 4))
        self.hilite_buttons[""] = clear_mark
        for color in NOTE_HILITES:
            button = Click(marks, " ", lambda picked=color: self._pick_highlight(picked), color, color, color, font=(FONT, 12), padx=8, pady=5, border=LINE)
            button.pack(side="left", padx=(0, 4))
            self.hilite_buttons[color] = button
        for key, label in (("left", "Sol"), ("center", "Orta"), ("right", "Sağ")):
            button = Click(marks, label, lambda picked=key: self._set_align(picked), CARD, INK, HOVER, font=(FONT, 12), padx=8, pady=5, border=LINE)
            button.pack(side="left", padx=(6 if key == "left" else 0, 4))
            self.align_buttons[key] = button
        lists = tk.Frame(card, bg=CARD)
        lists.pack(fill="x", pady=(6, 0))
        Click(lists, "Madde", self._bullet, CARD, INK, HOVER, font=(FONT, 12), padx=10, pady=5, border=LINE).pack(side="left", padx=(0, 4))
        Click(lists, "Numara", self._number, CARD, INK, HOVER, font=(FONT, 12), padx=10, pady=5, border=LINE).pack(side="left", padx=(0, 4))
        Click(lists, "İçeri", lambda: self._shift_indent(1), CARD, INK, HOVER, font=(FONT, 12), padx=10, pady=5, border=LINE).pack(side="left", padx=(0, 4))
        Click(lists, "Dışarı", lambda: self._shift_indent(-1), CARD, INK, HOVER, font=(FONT, 12), padx=10, pady=5, border=LINE).pack(side="left", padx=(0, 10))
        Click(lists, "Geri al", self._undo, CARD, INK, HOVER, font=(FONT, 12), padx=10, pady=5, border=LINE).pack(side="left", padx=(0, 4))
        Click(lists, "Yinele", self._redo, CARD, INK, HOVER, font=(FONT, 12), padx=10, pady=5, border=LINE).pack(side="left", padx=(0, 4))
        Click(lists, "Temizle", self._clear_format, CARD, INK, HOVER, font=(FONT, 12), padx=10, pady=5, border=LINE).pack(side="left")
        tools = tk.Frame(card, bg=CARD)
        tools.pack(fill="x", pady=(6, 0))
        for kind, label in (("rect", "Kutu"), ("oval", "Elips"), ("line", "Çizgi")):
            Click(tools, label, lambda picked=kind: self._insert_shape(picked), CARD, INK, HOVER, font=(FONT, 12), padx=10, pady=5, border=LINE).pack(side="left", padx=(0, 4))
        Click(tools, "Resim", self._add_image, CARD, INK, HOVER, font=(FONT, 12), padx=10, pady=5, border=LINE).pack(side="left")
        style = tk.Frame(card, bg=CARD)
        style.pack(fill="x", pady=(8, 0))
        tk.Label(style, text="Boy", bg=CARD, fg=MUTED, font=(FONT, 12)).pack(side="left", padx=(0, 6))
        for size in NOTE_SIZES:
            button = Click(style, str(size), lambda picked=size: self._pick_size(picked), CARD, INK, HOVER, font=(FONT, 12), padx=8, pady=4, border=LINE)
            button.pack(side="left", padx=(0, 4))
            self.size_buttons[size] = button
        tk.Label(style, text="Renk", bg=CARD, fg=MUTED, font=(FONT, 12)).pack(side="left", padx=(10, 6))
        for color in NOTE_COLORS:
            button = Click(style, "●", lambda picked=color: self._pick_color(picked), CARD, color, HOVER, font=(FONT, 16), padx=4, pady=0, border=LINE)
            button.pack(side="left", padx=(0, 4))
            self.color_buttons[color] = button
        stage = tk.Frame(card, bg=CARD)
        stage.pack(fill="both", expand=True, pady=(8, 0))
        self.paper = tk.Text(stage, width=78, height=16, wrap="word", undo=True, maxundo=80, font=(FONT, 18), fg=NOTE_COLORS[0], bg="#fffdf9", insertbackground=INK, relief="flat", padx=18, pady=16, highlightthickness=1, highlightbackground=LINE, spacing1=3, spacing3=6)
        scroll = ttk.Scrollbar(stage, orient="vertical", command=self.paper.yview)
        self.paper.configure(yscrollcommand=scroll.set)
        self.paper.pack(side="left")
        scroll.pack(side="left", fill="y")
        for size in NOTE_SIZES:
            for bold, italic in ((False, False), (True, False), (False, True), (True, True)):
                slant = "bold italic" if bold and italic else "bold" if bold else "italic" if italic else "normal"
                self.paper.tag_configure(self._font_tag(size, bold, italic), font=(FONT, size, slant))
        for color in NOTE_COLORS:
            self.paper.tag_configure(f"c{color[1:]}", foreground=color)
        for color in NOTE_HILITES:
            self.paper.tag_configure(f"hl{color[1:]}", background=color)
        self.paper.tag_configure("u", underline=True)
        self.paper.tag_configure("st", overstrike=True)
        for name in ("left", "center", "right"):
            self.paper.tag_configure(f"j{name}", justify=name)
        for level in range(1, 5):
            self.paper.tag_configure(f"in{level}", lmargin1=18 * level, lmargin2=18 * level)
        self.paper.tag_raise("u")
        self.paper.tag_raise("st")
        self.paper.bind("<KeyRelease>", self._stamp_last)
        self.paper.bind("<<Paste>>", self._paste)
        self.paper.bind("<Return>", self._on_return)
        self.paper.bind("<Command-b>", lambda _e: self._toggle("bold"))
        self.paper.bind("<Command-i>", lambda _e: self._toggle("italic"))
        self.paper.bind("<Command-u>", lambda _e: self._toggle("underline"))
        self.paper.bind("<ButtonRelease-1>", lambda _e: self._sync_marks())
        self.hint = tk.Label(card, text="K kalın, İ italik, A altı çizili, Ü üstü çizili. Seçili yazıya uygulanır.", bg=CARD, fg=MUTED, font=(FONT, 11))
        self.hint.pack(anchor="w", pady=(8, 0))
        buttons = tk.Frame(card, bg=CARD)
        buttons.pack(fill="x", pady=(10, 0))
        Click(buttons, "Vazgeç", self._cancel, CANVAS, INK, HOVER, font=(FONT, 12), padx=12, pady=6, border=LINE).pack(side="right")
        Click(buttons, "PDF", self._pdf, CARD, INK, HOVER, font=(FONT, 12), padx=12, pady=6, border=LINE).pack(side="right", padx=(0, 8))
        Click(buttons, "Kaydet", self._save, ACCENT, "white", ACCENT_HOVER, font=(FONT, 12, "bold"), padx=14, pady=6).pack(side="right", padx=(0, 8))
        self._show_size()
        self._show_color()
        self._show_marks()
        self._load(raw)
        self.bind("<Escape>", lambda _event: self._cancel())
        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self.geometry(f"+{parent.winfo_rootx() + 40}+{parent.winfo_rooty() + 30}")
        self.grab_set()
        self.paper.focus_set()

    def _font_tag(self, size, bold, italic):
        mark = ("b" if bold else "") + ("i" if italic else "")
        return f"sz{int(size)}{mark}"

    def _show_size(self):
        for key, button in self.size_buttons.items():
            if key == self.size:
                button.set_palette(ACCENT, "white", ACCENT_HOVER)
            else:
                button.set_palette(CARD, INK, HOVER, LINE)

    def _show_color(self):
        for key, button in self.color_buttons.items():
            button.set_palette(ACCENT if key == self.color else CARD, key, HOVER, ACCENT if key == self.color else LINE)

    def _show_marks(self):
        flags = {"bold": self.bold, "italic": self.italic, "underline": self.underline, "strike": self.strike}
        for key, button in self.mark_buttons.items():
            if flags[key]:
                button.set_palette(ACCENT, "white", ACCENT_HOVER)
            else:
                button.set_palette(CARD, INK, HOVER, LINE)
        for key, button in self.align_buttons.items():
            if key == self.align:
                button.set_palette(ACCENT, "white", ACCENT_HOVER)
            else:
                button.set_palette(CARD, INK, HOVER, LINE)
        for key, button in self.hilite_buttons.items():
            if key == self.highlight:
                button.set_palette(ACCENT if not key else key, "white" if not key else key, ACCENT_HOVER if not key else key, ACCENT)
            elif key:
                button.set_palette(key, key, key, LINE)
            else:
                button.set_palette(CARD, INK, HOVER, LINE)

    def _pick_size(self, size):
        self.size = size
        self._show_size()
        if self.paper.tag_ranges("sel"):
            self._retile("sel.first", "sel.last", lambda style: {**style, "size": size})
        self.paper.focus_set()

    def _pick_color(self, color):
        self.color = color
        self._show_color()
        if self.paper.tag_ranges("sel"):
            self._retile("sel.first", "sel.last", lambda style: {**style, "color": color})
        self.paper.focus_set()

    def _pick_highlight(self, color):
        self.highlight = color
        self._show_marks()
        if self.paper.tag_ranges("sel"):
            self._retile("sel.first", "sel.last", lambda style: {**style, "highlight": color})
        self.paper.focus_set()

    def _toggle(self, key):
        current = not getattr(self, key)
        setattr(self, key, current)
        self._show_marks()
        if self.paper.tag_ranges("sel"):
            self._retile("sel.first", "sel.last", lambda style, name=key, value=current: {**style, name: value})
        self.paper.focus_set()
        return "break"

    def _set_align(self, name):
        self.align = name
        self._show_marks()
        start = self.paper.index("sel.first linestart") if self.paper.tag_ranges("sel") else self.paper.index("insert linestart")
        end = self.paper.index("sel.last lineend") if self.paper.tag_ranges("sel") else self.paper.index("insert lineend")
        self._clear_style(start, end, "j")
        if name != "left":
            self.paper.tag_add(f"j{name}", start, end)
        self.paper.focus_set()

    def _shift_indent(self, step):
        start = self.paper.index("sel.first linestart") if self.paper.tag_ranges("sel") else self.paper.index("insert linestart")
        end = self.paper.index("sel.last lineend") if self.paper.tag_ranges("sel") else self.paper.index("insert lineend")
        level = 0
        for tag in self.paper.tag_names(start):
            if tag.startswith("in") and tag[2:].isdigit():
                level = int(tag[2:])
        level = max(0, min(4, level + step))
        self.indent = level
        self._clear_style(start, end, "in")
        if level:
            self.paper.tag_add(f"in{level}", start, end)
        self.paper.focus_set()

    def _bullet(self):
        self._prefix_line("• ")

    def _number(self):
        previous = self.paper.get("insert linestart -1 lines", "insert linestart")
        match = re.match(r"(\d+)\.\s", previous)
        number = int(match.group(1)) + 1 if match else 1
        self._prefix_line(f"{number}. ")

    def _prefix_line(self, prefix):
        line = self.paper.get("insert linestart", "insert lineend")
        if line.startswith(prefix):
            return
        start = self.paper.index("insert linestart")
        self.paper.edit_separator()
        self.paper.insert(start, prefix)
        self._paint(start, self.paper.index(f"{start}+{len(prefix)}c"))
        self.paper.focus_set()

    def _on_return(self, _event):
        line = self.paper.get("insert linestart", "insert lineend")
        bullet = line.startswith("• ")
        number = re.match(r"(\d+)\.\s", line)
        if bullet and line.strip() in {"•", "•"} or (number and re.fullmatch(r"\d+\.\s?", line.strip())):
            self.paper.delete("insert linestart", "insert lineend")
            return "break"
        if bullet:
            self.paper.insert("insert", "\n• ")
            self._paint("insert-3c", "insert")
            return "break"
        if number:
            nxt = int(number.group(1)) + 1
            text = f"\n{nxt}. "
            self.paper.insert("insert", text)
            self._paint(f"insert-{len(text)}c", "insert")
            return "break"
        return None

    def _clear_format(self):
        start = self.paper.index("sel.first") if self.paper.tag_ranges("sel") else self.paper.index("insert linestart")
        end = self.paper.index("sel.last") if self.paper.tag_ranges("sel") else self.paper.index("insert lineend")
        for prefix in ("sz", "c", "hl", "j", "in"):
            self._clear_style(start, end, prefix)
        self.paper.tag_remove("u", start, end)
        self.paper.tag_remove("st", start, end)
        self.bold = self.italic = self.underline = self.strike = False
        self.highlight = ""
        self.align = "left"
        self.indent = 0
        self._show_marks()
        self.paper.focus_set()

    def _style_at(self, index):
        names = set(self.paper.tag_names(index))
        style = {"size": self.size, "color": self.color, "bold": False, "italic": False, "underline": "u" in names, "strike": "st" in names, "highlight": "", "align": "left", "indent": 0}
        for tag in names:
            match = re.fullmatch(r"sz(\d+)([bi]*)", tag)
            if match and int(match.group(1)) in NOTE_SIZES:
                style["size"] = int(match.group(1))
                style["bold"] = "b" in match.group(2)
                style["italic"] = "i" in match.group(2)
            elif tag.startswith("c") and f"#{tag[1:]}" in NOTE_COLORS:
                style["color"] = f"#{tag[1:]}"
            elif tag.startswith("hl") and f"#{tag[2:]}" in NOTE_HILITES:
                style["highlight"] = f"#{tag[2:]}"
            elif tag in {"jleft", "jcenter", "jright"}:
                style["align"] = tag[1:]
            elif tag.startswith("in") and tag[2:].isdigit():
                style["indent"] = int(tag[2:])
        return style

    def _apply_style(self, start, end, style):
        for prefix in ("sz", "c", "hl"):
            self._clear_style(start, end, prefix)
        self.paper.tag_remove("u", start, end)
        self.paper.tag_remove("st", start, end)
        self.paper.tag_add(self._font_tag(style["size"], style["bold"], style["italic"]), start, end)
        self.paper.tag_add(f"c{style['color'][1:]}", start, end)
        if style.get("underline"):
            self.paper.tag_add("u", start, end)
        if style.get("strike"):
            self.paper.tag_add("st", start, end)
        if style.get("highlight"):
            self.paper.tag_add(f"hl{style['highlight'][1:]}", start, end)

    def _retile(self, start, end, change):
        start = self.paper.index(start)
        end = self.paper.index(end)
        pieces = []
        index = start
        while self.paper.compare(index, "<", end):
            style = self._style_at(index)
            nxt = self.paper.index(f"{index}+1c")
            pieces.append((index, nxt, change(style)))
            index = nxt
        merged = []
        for piece in pieces:
            if merged and merged[-1][2] == piece[2] and merged[-1][1] == piece[0]:
                merged[-1] = (merged[-1][0], piece[1], piece[2])
            else:
                merged.append(piece)
        for begin, stop, style in merged:
            self._apply_style(begin, stop, style)

    def _clear_style(self, start, end, prefix):
        for tag in self.paper.tag_names():
            if tag.startswith(prefix):
                self.paper.tag_remove(tag, start, end)

    def _paint(self, start, end):
        if self.paper.compare(start, ">=", end):
            return
        self._apply_style(start, end, {
            "size": self.size,
            "color": self.color,
            "bold": self.bold,
            "italic": self.italic,
            "underline": self.underline,
            "strike": self.strike,
            "highlight": self.highlight,
        })

    def _stamp_last(self, event):
        if event.keysym in {"BackSpace", "Delete", "Left", "Right", "Up", "Down", "Home", "End", "Prior", "Next", "Tab"} or not event.char:
            self._sync_marks()
            return
        self._paint("insert-1c", "insert")
        self._sync_marks()

    def _sync_marks(self):
        if self.paper.compare("insert", "==", "1.0"):
            return
        style = self._style_at("insert-1c")
        self.size = style["size"]
        self.color = style["color"]
        self.bold = style["bold"]
        self.italic = style["italic"]
        self.underline = style["underline"]
        self.strike = style["strike"]
        self.highlight = style["highlight"]
        self.align = style["align"]
        self.indent = style["indent"]
        self._show_size()
        self._show_color()
        self._show_marks()

    def _paste(self, _event):
        try:
            clip = self.clipboard_get()
        except tk.TclError:
            return "break"
        clip = clip[:4000]
        start = self.paper.index("insert")
        self.paper.edit_separator()
        self.paper.insert("insert", clip)
        self._paint(start, "insert")
        return "break"

    def _break_line(self):
        if self.paper.compare("insert", ">", "insert linestart"):
            self.paper.insert("insert", "\n")

    def _insert_shape(self, kind):
        self._break_line()
        width, height = (180, 28) if kind == "line" else (180, 72)
        box = tk.Canvas(self.paper, width=width, height=height, bg="#fffdf9", highlightthickness=0)
        if kind == "oval":
            box.create_oval(6, 6, width - 6, height - 6, outline=self.color, width=2)
        elif kind == "line":
            box.create_line(6, height / 2, width - 6, height / 2, fill=self.color, width=2)
        else:
            box.create_rectangle(6, 6, width - 6, height - 6, outline=self.color, width=2)
        self.paper.window_create("insert", window=box)
        self.shapes[str(box)] = {"type": "shape", "kind": kind, "w": width, "h": height, "color": self.color}
        self.paper.insert("insert", "\n")
        self.paper.focus_set()

    def _add_image(self):
        selected = filedialog.askopenfilename(
            parent=self,
            title="Görsel ekle",
            filetypes=[("Görsel", "*.png *.jpg *.jpeg *.gif *.webp *.tif *.tiff *.heic")],
        )
        if not selected:
            return
        try:
            dest = store_note_image(self.task_id, selected)
            photo, width, height = fit_note_photo(dest)
        except (OSError, ValueError, tk.TclError) as exc:
            messagebox.showwarning(APP_NAME, str(exc) or "Görsel açılamadı.", parent=self)
            return
        self.photos.append(photo)
        self._break_line()
        self.paper.image_create("insert", image=photo)
        self.images[str(photo)] = {"type": "image", "path": str(dest), "w": width, "h": height}
        self.paper.insert("insert", "\n")
        self.paper.focus_set()

    def _undo(self):
        try:
            self.paper.edit_undo()
        except tk.TclError:
            return
        self.paper.focus_set()

    def _redo(self):
        try:
            self.paper.edit_redo()
        except tk.TclError:
            return
        self.paper.focus_set()

    def _load(self, raw):
        data = _note_data(raw)
        blocks = data.get("blocks") if isinstance(data.get("blocks"), list) else None
        if blocks is None:
            blocks = []
            for item in load_note_items(raw):
                if item["type"] == "text":
                    blocks.append({"type": "text", "text": item["text"], "size": item["size"], "color": item["color"]})
                elif item["type"] == "image":
                    blocks.append({"type": "image", "path": item["path"], "w": item["w"], "h": item["h"]})
                elif item["type"] == "shape":
                    blocks.append({"type": "shape", "kind": item["kind"], "w": 180, "h": 72, "color": item["color"]})
        self.paper.edit_separator()
        for block in blocks[:200]:
            if not isinstance(block, dict):
                continue
            if block.get("type") == "text":
                text = str(block.get("text") or "")[:4000]
                if not text:
                    continue
                start = self.paper.index("insert")
                self.paper.insert("insert", text)
                size = int(block.get("size") or 18)
                self.size = size if size in NOTE_SIZES else 18
                color = block.get("color") if block.get("color") in NOTE_COLORS else NOTE_COLORS[0]
                self.color = color
                self.bold = bool(block.get("bold"))
                self.italic = bool(block.get("italic"))
                self.underline = bool(block.get("underline"))
                self.strike = bool(block.get("strike"))
                mark = block.get("highlight") if block.get("highlight") in NOTE_HILITES else ""
                self.highlight = mark
                self._paint(start, "insert")
                align = block.get("align") if block.get("align") in {"center", "right"} else "left"
                if align != "left":
                    self.paper.tag_add(f"j{align}", start, "insert")
                try:
                    indent = int(block.get("indent") or 0)
                except (TypeError, ValueError):
                    indent = 0
                if 1 <= indent <= 4:
                    self.paper.tag_add(f"in{indent}", start, "insert")
            elif block.get("type") == "image":
                path = Path(str(block.get("path") or ""))
                if not path.is_file() or path.is_symlink() or not is_within(path, NOTE_DIR):
                    continue
                try:
                    photo, width, height = fit_note_photo(path)
                except tk.TclError:
                    continue
                self.photos.append(photo)
                self.paper.image_create("insert", image=photo)
                self.images[str(photo)] = {"type": "image", "path": str(path.resolve()), "w": width, "h": height}
            elif block.get("type") == "shape" and block.get("kind") in {"rect", "oval", "line"}:
                self.color = block.get("color") if block.get("color") in NOTE_COLORS else NOTE_COLORS[0]
                self._insert_shape(block["kind"])
        self._show_size()
        self._show_color()
        self._show_marks()
        self.paper.edit_reset()
        self.paper.mark_set("insert", "1.0")

    def _style_from_tags(self, names):
        style = {"size": self.size, "color": self.color, "bold": False, "italic": False, "underline": "u" in names, "strike": "st" in names, "highlight": "", "align": "left", "indent": 0}
        for tag in names:
            match = re.fullmatch(r"sz(\d+)([bi]*)", tag)
            if match and int(match.group(1)) in NOTE_SIZES:
                style["size"] = int(match.group(1))
                style["bold"] = "b" in match.group(2)
                style["italic"] = "i" in match.group(2)
            elif tag.startswith("c") and f"#{tag[1:]}" in NOTE_COLORS:
                style["color"] = f"#{tag[1:]}"
            elif tag.startswith("hl") and f"#{tag[2:]}" in NOTE_HILITES:
                style["highlight"] = f"#{tag[2:]}"
            elif tag in {"jleft", "jcenter", "jright"}:
                style["align"] = tag[1:]
            elif tag.startswith("in") and tag[2:].isdigit():
                style["indent"] = int(tag[2:])
        return style

    def _document(self):
        blocks = []
        buffer = []
        active = None
        live = set()

        def flush():
            nonlocal buffer, active
            if not buffer:
                return
            block = {"type": "text", "text": "".join(buffer), "size": active["size"], "color": active["color"]}
            if active["bold"]:
                block["bold"] = True
            if active["italic"]:
                block["italic"] = True
            if active["underline"]:
                block["underline"] = True
            if active["strike"]:
                block["strike"] = True
            if active["highlight"]:
                block["highlight"] = active["highlight"]
            if active["align"] != "left":
                block["align"] = active["align"]
            if active["indent"]:
                block["indent"] = active["indent"]
            blocks.append(block)
            buffer = []

        for key, value, _index in self.paper.dump("1.0", "end-1c", text=True, image=True, window=True, tag=True):
            if key == "tagon":
                live.add(value)
            elif key == "tagoff":
                live.discard(value)
            elif key == "text":
                style = self._style_from_tags(live)
                if buffer and style != active:
                    flush()
                active = style
                buffer.append(value)
            elif key == "image":
                flush()
                info = self.images.get(value)
                if info:
                    blocks.append(dict(info))
            elif key == "window":
                flush()
                info = self.shapes.get(value)
                if info:
                    blocks.append(dict(info))
        flush()
        body = self.paper.get("1.0", "end-1c")
        return {"body": body, "blocks": blocks[:200]}

    def _pdf(self):
        document = self._document()
        useful = [block for block in document["blocks"] if block.get("type") != "text" or str(block.get("text") or "").strip()]
        if not useful:
            messagebox.showinfo(APP_NAME, "PDF için nota bir şey yaz.", parent=self)
            return
        path = filedialog.asksaveasfilename(
            parent=self,
            title="Notu PDF kaydet",
            defaultextension=".pdf",
            initialfile=f"{safe_name(self.task_title)}.pdf",
            filetypes=[("PDF", "*.pdf")],
        )
        if not path:
            return
        try:
            write_note_pdf(path, self.task_title, document["blocks"])
        except (OSError, ValueError) as exc:
            messagebox.showwarning(APP_NAME, str(exc), parent=self)
            return
        self.keep = True
        self.hint.configure(text="PDF kaydedildi. Not görevde de durur.")

    def _save(self):
        self.keep = True
        self.saved = json.dumps(self._document(), ensure_ascii=False)
        self.destroy()

    def _cancel(self):
        self.saved = json.dumps(self._document(), ensure_ascii=False) if self.keep else None
        self.destroy()


class InvoicePreview(tk.Frame):
    def __init__(self, master, app, detail=False):
        super().__init__(master, bg=CARD, highlightbackground=LINE, highlightthickness=1, width=420 if detail else 372)
        self.app = app
        self.detail = detail
        self.row = None
        self.photo = None
        self.loaded = None
        self.generation = 0
        self.summary = None
        self.extra = None
        self.remain_value = None
        if detail:
            self.footer = tk.Frame(self, bg=CARD)
            self.footer.pack(side="bottom", fill="x")
            tk.Frame(self.footer, bg=LINE, height=1).pack(fill="x")
            foot = tk.Frame(self.footer, bg=CARD)
            foot.pack(fill="x", padx=14, pady=10)
            captions = tk.Frame(foot, bg=CARD)
            captions.pack(side="left", fill="x", expand=True)
            tk.Label(captions, text="Bekleyen tahsilat", bg=CARD, fg=MUTED, font=(FONT, 11)).pack(anchor="w")
            self.remain_value = tk.Label(captions, text="—", bg=CARD, fg=INK, font=(FONT, 16, "bold"))
            self.remain_value.pack(anchor="w")
            Click(foot, "Tahsil et", self._collect, ACCENT, "white", ACCENT_HOVER, font=(FONT, 12, "bold"), padx=12, pady=8).pack(side="right")
            holder = tk.Frame(self, bg=CARD)
            holder.pack(fill="both", expand=True)
            self.scroll = tk.Canvas(holder, bg=CARD, highlightthickness=0, width=400)
            scroll = ttk.Scrollbar(holder, orient="vertical", command=self.scroll.yview)
            self.scroll.configure(yscrollcommand=scroll.set)
            self.scroll.pack(side="left", fill="both", expand=True)
            scroll.pack(side="right", fill="y")
            self.body = tk.Frame(self.scroll, bg=CARD)
            self._body_window = self.scroll.create_window((0, 0), window=self.body, anchor="nw")
            self.body.bind("<Configure>", lambda _e: self.scroll.configure(scrollregion=self.scroll.bbox("all")))
            self.scroll.bind("<Configure>", lambda event: self.scroll.itemconfigure(self._body_window, width=event.width))
            self.scroll.bind("<MouseWheel>", self._wheel)
            self.body.bind("<MouseWheel>", self._wheel)
            parent = self.body
            self.summary = tk.Frame(parent, bg=CARD)
            self.summary.pack(fill="x", padx=12, pady=(12, 0))
        else:
            parent = self
        tk.Label(parent, text="Fatura", bg=CARD, fg=ACCENT, font=(FONT, 12, "bold")).pack(anchor="w", padx=16, pady=(16, 0))
        self.heading = tk.Label(parent, text="Henüz fatura yok", bg=CARD, fg=INK, font=(FONT, 16, "bold"), wraplength=320, justify="left")
        self.heading.pack(anchor="w", padx=16, pady=(4, 0))
        self.meta = tk.Label(parent, text="Seçilen fatura burada açılır.", bg=CARD, fg=MUTED, font=(FONT, 12), wraplength=320, justify="left")
        self.meta.pack(anchor="w", padx=16, pady=(2, 8))
        self.stage = tk.Frame(parent, bg=PAPER, highlightbackground=LINE, highlightthickness=1, height=210)
        self.stage.pack(fill="x", padx=16, pady=(0, 8))
        self.stage.pack_propagate(False)
        self.image = tk.Label(self.stage, text="Önizleme\nburada", bg=PAPER, fg=MUTED, font=(FONT, 13), cursor="hand2", justify="center")
        self.image.pack(fill="both", expand=True)
        self.image.bind("<Button-1>", lambda _e: self.enlarge())
        self.stage.bind("<Button-1>", lambda _e: self.enlarge())
        self.caption = tk.Label(parent, text="Görsele basınca büyür.", bg=CARD, fg=MUTED, font=(FONT, 11))
        self.caption.pack(anchor="w", padx=16)
        actions = tk.Frame(parent, bg=CARD)
        actions.pack(fill="x", padx=16, pady=(8, 4))
        Click(actions, "Büyüt", self.enlarge, ACCENT, "white", ACCENT_HOVER, font=(FONT, 12, "bold"), padx=12, pady=6).pack(side="left")
        if detail:
            Click(actions, "Dosyayı aç", self.open_file, CARD, INK, HOVER, font=(FONT, 12), padx=12, pady=6, border=LINE).pack(side="left", padx=(8, 0))
            self.extra = tk.Frame(parent, bg=CARD)
            self.extra.pack(fill="x", padx=12, pady=(4, 16))

    def _wheel(self, event):
        if self.detail:
            self.scroll.yview_scroll(-1 if event.delta > 0 else 1, "units")

    def show(self, row):
        self.row = row
        if self.detail:
            self._paint_detail(row)
        if row is None:
            self.heading.configure(text="Henüz fatura yok")
            self.meta.configure(text="Seçilen fatura burada açılır.")
            self._set_image(None, "Önizleme\nburada")
            self.loaded = None
            return
        amount = format_amount(row["amount"])
        amount_text = f"{amount} ₺" if amount else "Tutar bekleniyor"
        source = {"klasor": "Klasör", "eposta": "E-posta", "manuel": "Elle"}.get(row["source"] or "manuel", "Elle")
        self.heading.configure(text=row["party"] or "Fatura")
        self.meta.configure(text=f"{amount_text}   ·   {row['invoice_date'] or 'tarih yok'}   ·   {source}\n{row['excel_status'] or row['status'] or ''}")
        attachment = row["attachment"] or ""
        try:
            path = assert_openable(attachment, self.app.allowed_roots()) if attachment else None
        except ValueError as exc:
            self._set_image(None, str(exc))
            self.loaded = None
            return
        if path is None:
            self._set_image(None, "Bu kayıtta görüntü yok")
            self.loaded = None
            return
        if self.loaded == path and self.photo is not None:
            self.caption.configure(text="Büyütmek için görsele bas.")
            return
        self.generation += 1
        generation = self.generation
        self.caption.configure(text="Görüntü hazırlanıyor…")
        threading.Thread(target=self._render, args=(path, generation), daemon=True).start()

    def _render(self, path, generation):
        try:
            image = render_preview(path, 900)
        except (OSError, subprocess.SubprocessError):
            image = None
        self.app.events.put(("preview", id(self), generation, str(image) if image else ""))

    def receive(self, generation, image_path):
        if generation != self.generation:
            return
        if not image_path:
            self._set_image(None, "Önizleme üretilemedi")
            self.caption.configure(text="Dosyayı yine de Önizleme ile açabilirsin.")
            return
        try:
            self.photo = fit_photo(image_path, 328, 190)
        except tk.TclError:
            self._set_image(None, "Görüntü okunamadı")
            return
        self.image.configure(image=self.photo, text="")
        self.loaded = Path(self.row["attachment"]).resolve() if self.row and self.row["attachment"] else None
        self.caption.configure(text="Büyütmek için görsele bas.")

    def _set_image(self, photo, text):
        self.photo = photo
        self.image.configure(image=photo or "", text=text)

    def enlarge(self):
        if not self.row or not self.row["attachment"]:
            return
        try:
            path = assert_openable(self.row["attachment"], self.app.allowed_roots())
        except ValueError as exc:
            messagebox.showwarning(APP_NAME, str(exc), parent=self.app)
            return
        self.caption.configure(text="Büyük görüntü hazırlanıyor…")
        longest = max(self.app.winfo_screenwidth(), self.app.winfo_screenheight(), 1600) * 2
        threading.Thread(target=self._render_large, args=(path, self.generation + 1, min(longest, 3200)), daemon=True).start()
        self.generation += 1

    def _render_large(self, path, generation, size):
        try:
            image = render_preview(path, size)
        except (OSError, subprocess.SubprocessError):
            image = None
        self.app.events.put(("lightbox", id(self), generation, str(image) if image else "", self.row["party"] if self.row else "Fatura"))

    def open_lightbox(self, image_path, title):
        self.caption.configure(text="Büyütmek için görsele bas.")
        if not image_path:
            messagebox.showinfo(APP_NAME, "Büyük görüntü üretilemedi.", parent=self.app)
            return
        window = tk.Toplevel(self.app)
        window.title(title or "Fatura")
        window.configure(bg="#12161c")
        try:
            window.attributes("-fullscreen", True)
        except tk.TclError:
            window.geometry(f"{window.winfo_screenwidth()}x{window.winfo_screenheight()}+0+0")
        window.update_idletasks()
        max_w = max(window.winfo_screenwidth() - 48, 800)
        max_h = max(window.winfo_screenheight() - 72, 600)
        try:
            photo = fit_photo(image_path, max_w, max_h)
        except tk.TclError:
            window.destroy()
            return
        window.photo = photo
        label = tk.Label(window, image=photo, bg="#12161c", cursor="hand2")
        label.pack(expand=True)
        hint = tk.Label(window, text="Kapatmak için görsele veya Esc tuşuna bas", bg="#12161c", fg="#c5ced8", font=(FONT, 13))
        hint.pack(pady=(0, 18))
        label.bind("<Button-1>", lambda _e: window.destroy())
        window.bind("<Escape>", lambda _e: window.destroy())
        window.lift()
        window.focus_force()
        window.grab_set()

    def open_file(self):
        if not self.row or not self.row["attachment"]:
            return
        try:
            path = assert_openable(self.row["attachment"], self.app.allowed_roots())
        except ValueError as exc:
            messagebox.showwarning(APP_NAME, str(exc), parent=self.app)
            return
        open_path(path)

    def _collect(self):
        if self.row:
            self.app.collect_invoice(self.row["id"])

    def _paint_detail(self, row):
        if self.summary is None:
            return
        for child in self.summary.winfo_children():
            child.destroy()
        for child in self.extra.winfo_children():
            child.destroy()
        if row is None:
            if self.remain_value is not None:
                self.remain_value.configure(text="—")
            return
        total, paid = invoice_money(row)
        remaining = max(total - paid, 0)
        ratio = (paid / total) if total else 0
        status = row["status"] or "Ödenmedi"
        pill = STATUS_PILL.get(status, (PAPER, MUTED))
        card = tk.Frame(self.summary, bg=CARD, highlightbackground=LINE, highlightthickness=1)
        card.pack(fill="x")
        head = tk.Frame(card, bg=CARD)
        head.pack(fill="x", padx=14, pady=(12, 0))
        text = tk.Frame(head, bg=CARD)
        text.pack(side="left", fill="x", expand=True)
        tk.Label(text, text="Cari / müşteri", bg=CARD, fg=MUTED, font=(FONT, 10, "bold")).pack(anchor="w")
        tk.Label(text, text=row["party"] or "Fatura", bg=CARD, fg=INK, font=(FONT, 16, "bold"), wraplength=250, justify="left", anchor="w").pack(anchor="w")
        kind = row["kind"] or "Fatura"
        tk.Label(text, text=kind, bg=CARD, fg=MUTED, font=(FONT, 11), anchor="w").pack(anchor="w")
        tk.Label(head, text=party_initials(row["party"]), bg=SOFT, fg=ACCENT, font=(FONT, 12, "bold"), width=4, pady=8).pack(side="right")
        tk.Frame(card, bg=LINE, height=1).pack(fill="x", padx=14, pady=10)
        money = tk.Frame(card, bg=CARD)
        money.pack(fill="x", padx=14)
        tk.Label(money, text="Genel toplam", bg=CARD, fg=MUTED, font=(FONT, 10, "bold")).pack(anchor="w")
        amount = format_amount(total)
        tk.Label(money, text=(amount + " ₺") if amount else "Tutar yok", bg=CARD, fg=INK, font=(FONT, 22, "bold")).pack(side="left")
        tk.Label(money, text=status, bg=pill[0], fg=pill[1], font=(FONT, 10, "bold"), padx=8, pady=3).pack(side="right")
        due = parse_day(row["due_date"])
        if status in ("Ödenmedi", "Kısmi ödendi") and due:
            delta = (due - date.today()).days
            if delta < 0:
                tone, note = (DANGER_SOFT, DANGER), f"{abs(delta)} gün gecikti"
            elif delta == 0:
                tone, note = (WARN_SOFT, WARN), "Bugün"
            else:
                tone, note = (WARN_SOFT, WARN), f"{delta} gün kaldı"
            banner = tk.Frame(card, bg=tone[0])
            banner.pack(fill="x", padx=14, pady=(10, 0))
            tk.Label(banner, text=f"Son ödeme: {due.day} {MONTHS[due.month - 1]} {due.year}", bg=tone[0], fg=tone[1], font=(FONT, 11)).pack(side="left", padx=8, pady=6)
            tk.Label(banner, text=note, bg=tone[0], fg=tone[1], font=(FONT, 11, "bold")).pack(side="right", padx=8)
        track = tk.Frame(card, bg=CARD)
        track.pack(fill="x", padx=14, pady=(10, 12))
        labels = tk.Frame(track, bg=CARD)
        labels.pack(fill="x")
        percent = f"%{round(ratio * 100)}"
        tk.Label(labels, text=f"Tahsilat {percent}", bg=CARD, fg=MUTED, font=(FONT, 11)).pack(side="left")
        left = format_amount(remaining)
        tk.Label(labels, text=("Kalan " + left + " ₺") if left else "Kalan yok", bg=CARD, fg=INK, font=(FONT, 11, "bold")).pack(side="right")
        bar = tk.Canvas(track, height=8, bg=PAPER, highlightthickness=0)
        bar.pack(fill="x", pady=(6, 0))

        def draw(event, canvas=bar, share=ratio):
            canvas.delete("all")
            width = max(event.width, 1)
            canvas.create_rectangle(0, 0, width, 8, fill=PAPER, outline="")
            if share > 0:
                canvas.create_rectangle(0, 0, max(6, width * share), 8, fill=OK if share >= 1 else ACCENT, outline="")

        bar.bind("<Configure>", draw)
        if self.remain_value is not None:
            self.remain_value.configure(text=(format_amount(remaining) + " ₺") if total else "—", fg=OK if remaining <= 0 and total else INK)
        self._paint_timeline(row, total, paid)
        self._paint_actions(row)

    def _paint_timeline(self, row, total, paid):
        box = tk.Frame(self.extra, bg=CARD, highlightbackground=LINE, highlightthickness=1)
        box.pack(fill="x", pady=(0, 8))
        tk.Label(box, text="Tahsilat takibi", bg=CARD, fg=INK, font=(FONT, 13, "bold")).pack(anchor="w", padx=12, pady=(10, 6))
        steps = []
        created = str(row["created_at"] or "")[:16].replace("T", " ")
        steps.append((OK, "Kayıt oluştu", created or "tarih yok"))
        if row["invoice_date"]:
            steps.append((ACCENT, "Belge tarihi", row["invoice_date"]))
        source = {"klasor": "Klasörden alındı", "eposta": "E-postadan alındı"}.get(row["source"] or "", "Elle eklendi")
        steps.append((ACCENT, source, row["invoice_no"] or ""))
        for event in invoice_events(row):
            steps.append((OK if event.get("kind") == "payment" else ACCENT, event.get("text") or "İşlem", event.get("at") or ""))
        due = parse_day(row["due_date"])
        if row["status"] == "Ödendi":
            steps.append((OK, "Tahsilat tamam", format_amount(paid) + " ₺" if paid else ""))
        elif row["status"] == "İptal":
            steps.append((MUTED, "İptal edildi", ""))
        elif due:
            delta = (due - date.today()).days
            when = f"{due.day} {MONTHS[due.month - 1]} {due.year}"
            if delta < 0:
                steps.append((DANGER, "Vade geçti", when))
            else:
                steps.append((WARN, "Vade sonu tahsilat", when))
        elif total and paid < total:
            steps.append((WARN, "Tahsilat bekleniyor", format_amount(total - paid) + " ₺"))
        for color, title, note in steps:
            line = tk.Frame(box, bg=CARD)
            line.pack(fill="x", padx=12, pady=3)
            tk.Label(line, text="●", bg=CARD, fg=color, font=(FONT, 12)).pack(side="left", anchor="n")
            copy = tk.Frame(line, bg=CARD)
            copy.pack(side="left", fill="x", expand=True, padx=(6, 0))
            tk.Label(copy, text=title, bg=CARD, fg=INK, font=(FONT, 12, "bold"), anchor="w").pack(fill="x")
            if note:
                tk.Label(copy, text=note, bg=CARD, fg=MUTED, font=(FONT, 11), anchor="w").pack(fill="x")
        tk.Frame(box, bg=CARD, height=8).pack()

    def _paint_actions(self, row):
        box = tk.Frame(self.extra, bg=CARD, highlightbackground=LINE, highlightthickness=1)
        box.pack(fill="x")
        tk.Label(box, text="Hızlı işlem", bg=CARD, fg=MUTED, font=(FONT, 10, "bold")).pack(anchor="w", padx=12, pady=(10, 6))
        row_buttons = tk.Frame(box, bg=CARD)
        row_buttons.pack(fill="x", padx=12)
        Click(row_buttons, "WhatsApp", lambda: self.app.reach_invoice_party(row["id"], "whatsapp"), OK_SOFT, OK, "#d7f5e4", font=(FONT, 12, "bold"), padx=10, pady=6).pack(side="left", fill="x", expand=True)
        Click(row_buttons, "SMS", lambda: self.app.reach_invoice_party(row["id"], "sms"), CARD, INK, HOVER, font=(FONT, 12, "bold"), padx=10, pady=6, border=LINE).pack(side="left", fill="x", expand=True, padx=(8, 0))
        if row["status"] != "İptal":
            Click(box, "Faturayı iptal et", lambda: self.app.cancel_invoice_record(row["id"]), CARD, DANGER, DANGER_SOFT, font=(FONT, 12), padx=10, pady=6).pack(anchor="w", padx=12, pady=(8, 12))
        else:
            tk.Frame(box, bg=CARD, height=10).pack()
        self._bind_detail_wheel(self.summary)
        self._bind_detail_wheel(self.extra)

    def _bind_detail_wheel(self, widget):
        widget.bind("<MouseWheel>", self._wheel)
        for child in widget.winfo_children():
            self._bind_detail_wheel(child)


class OfferCompareDialog(tk.Toplevel):
    MATCH_COLORS = {
        "same": (OK_SOFT, OK),
        "similar": (WARN_SOFT, WARN),
        "different": (DANGER_SOFT, DANGER),
        "unknown": (PAPER, MUTED),
    }

    def __init__(self, app, project, rows):
        super().__init__(app)
        self.app = app
        self.project = project
        self.rows = list(rows)
        self.mode = "Dengeli"
        self.title(f"Teklif karşılaştırması · {project['name']}")
        self.configure(bg=CANVAS)
        self.transient(app)
        self.geometry(f"1180x820+{app.winfo_rootx() + 40}+{app.winfo_rooty() + 30}")
        self.minsize(960, 640)
        self._build_header()
        holder = tk.Frame(self, bg=CANVAS)
        holder.pack(fill="both", expand=True)
        self.canvas = tk.Canvas(holder, bg=CANVAS, highlightthickness=0)
        scroll = ttk.Scrollbar(holder, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=scroll.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        self.body = tk.Frame(self.canvas, bg=CANVAS)
        self._window = self.canvas.create_window((0, 0), window=self.body, anchor="nw")
        self.body.bind("<Configure>", lambda _e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", self._on_resize)
        self.bind_all("<MouseWheel>", self._on_wheel, add="+")
        self.bind("<Escape>", lambda _e: self.destroy())
        self._width = 1140
        self.render()

    def destroy(self):
        try:
            self.unbind_all("<MouseWheel>")
        except tk.TclError:
            pass
        super().destroy()

    def _on_wheel(self, event):
        if not self.winfo_exists():
            return
        try:
            if str(event.widget.winfo_toplevel()) != str(self):
                return
        except (tk.TclError, AttributeError, KeyError):
            return
        self.canvas.yview_scroll(-1 if event.delta > 0 else 1, "units")

    def _on_resize(self, event):
        self.canvas.itemconfigure(self._window, width=event.width)
        if abs(event.width - self._width) > 40:
            self._width = event.width
            self.after_idle(self.render)

    def _build_header(self):
        top = tk.Frame(self, bg=NAV)
        top.pack(fill="x")
        left = tk.Frame(top, bg=NAV)
        left.pack(side="left", padx=28, pady=(18, 16))
        tk.Label(left, text="TEKLİF KARŞILAŞTIRMASI", bg=NAV, fg=NAV_MUTED, font=(FONT, 11, "bold")).pack(anchor="w")
        tk.Label(left, text=self.project["name"], bg=NAV, fg="white", font=(FONT, 24, "bold")).pack(anchor="w")
        bits = [f"{len(self.rows)} teklif"]
        if self.project["budget"] is not None:
            bits.append(f"Bütçe {format_amount(self.project['budget'])} ₺")
        if self.project["end_date"]:
            bits.append(f"Bitiş {self.project['end_date']}")
        bits.append(f"Hazırlanma {long_today()}")
        tk.Label(left, text="   ·   ".join(bits), bg=NAV, fg=NAV_TEXT, font=(FONT, 12)).pack(anchor="w", pady=(2, 0))
        right = tk.Frame(top, bg=NAV)
        right.pack(side="right", padx=28)
        tk.Label(right, text="Öncelik", bg=NAV, fg=NAV_MUTED, font=(FONT, 11, "bold")).pack(anchor="e", pady=(0, 4))
        chips = tk.Frame(right, bg=NAV)
        chips.pack(anchor="e")
        self.mode_chips = {}
        for name in COMPARE_MODES:
            chip = Click(chips, name, lambda value=name: self._set_mode(value), NAV_ON, "white", NAV_HOVER, font=(FONT, 12), padx=12, pady=6)
            chip.pack(side="left", padx=(6, 0))
            self.mode_chips[name] = chip
        self._paint_modes()

    def _paint_modes(self):
        for name, chip in self.mode_chips.items():
            if name == self.mode:
                chip.set_palette(ACCENT, "white", ACCENT_HOVER)
            else:
                chip.set_palette(NAV_ON, NAV_TEXT, NAV_HOVER)

    def _set_mode(self, name):
        self.mode = name
        self._paint_modes()
        self.render()

    def _section(self, title, note=""):
        frame = tk.Frame(self.body, bg=CANVAS)
        frame.pack(fill="x", padx=28, pady=(22, 0))
        tk.Label(frame, text=title, bg=CANVAS, fg=INK, font=(FONT, 15, "bold"), anchor="w").pack(side="left")
        if note:
            tk.Label(frame, text=note, bg=CANVAS, fg=MUTED, font=(FONT, 11), anchor="w").pack(side="left", padx=(10, 0), pady=(3, 0))
        return frame

    def render(self):
        for child in self.body.winfo_children():
            child.destroy()
        data = compare_offers(self.project, self.rows, self.mode)
        self.data = data
        width = max(900, self.canvas.winfo_width() or self._width)
        self._decision(data, width)
        self._kpis(data)
        self._match(data, width)
        self._section("Puan ve fiyat / fayda", self._weights_note(data))
        self._score_cards(data, width)
        self._section("Yan yana kıyas", "Yeşil: o satırda en iyi. Kırmızı: en zayıf.")
        self._matrix(data, width)
        self._section("Karar öncesi netleştirilmesi gerekenler")
        self._gaps(data, width)
        tk.Frame(self.body, bg=CANVAS, height=24).pack(fill="x")
        foot = tk.Frame(self.body, bg=CANVAS)
        foot.pack(fill="x", padx=28, pady=(0, 24))
        Click(foot, "Kapat", self.destroy, CARD, INK, HOVER, font=(FONT, 12), padx=14, pady=6, border=LINE).pack(side="right")
        self.canvas.yview_moveto(0)

    def _weights_note(self, data):
        if not data["weights"]:
            return ""
        parts = [f"{COMPARE_LABELS[key]} %{round(value * 100)}" for key, value in data["weights"].items()]
        return "Ağırlık: " + ", ".join(parts)

    def _decision(self, data, width):
        pick = data["verdict"]["pick"]
        card = tk.Frame(self.body, bg=CARD, highlightbackground=LINE, highlightthickness=1)
        card.pack(fill="x", padx=28, pady=(22, 0))
        tk.Frame(card, bg=ACCENT, width=6).pack(side="left", fill="y")
        inner = tk.Frame(card, bg=CARD)
        inner.pack(side="left", fill="both", expand=True, padx=22, pady=18)
        score_box = tk.Frame(inner, bg=CARD)
        score_box.pack(side="right", padx=(16, 0))
        if pick is not None:
            tk.Label(score_box, text=str(pick["score"]), bg=CARD, fg=ACCENT, font=(FONT, 40, "bold")).pack(anchor="e")
            tk.Label(score_box, text="/ 100 puan", bg=CARD, fg=MUTED, font=(FONT, 11)).pack(anchor="e")
        text_box = tk.Frame(inner, bg=CARD)
        text_box.pack(side="left", fill="both", expand=True)
        tk.Label(text_box, text="ÖNERİ", bg=CARD, fg=ACCENT, font=(FONT, 11, "bold"), anchor="w").pack(fill="x")
        if pick is None:
            tk.Label(text_box, text="Karşılaştırılacak teklif yok.", bg=CARD, fg=INK, font=(FONT, 18, "bold"), anchor="w").pack(fill="x")
            return
        line = tk.Frame(text_box, bg=CARD)
        line.pack(fill="x", pady=(2, 0))
        tk.Label(line, text=pick["label"], bg=CARD, fg=INK, font=(FONT, 24, "bold")).pack(side="left")
        if pick["effective"] is not None:
            tk.Label(line, text=f"{format_amount(pick['effective'])} ₺", bg=CARD, fg=INK, font=(FONT, 18)).pack(side="left", padx=(14, 0), pady=(6, 0))
        tk.Label(text_box, text=data["verdict"]["text"], bg=CARD, fg=INK, font=(FONT, 13), anchor="w", justify="left", wraplength=width - 260).pack(fill="x", pady=(8, 0))

    def _kpis(self, data):
        row = tk.Frame(self.body, bg=CANVAS)
        row.pack(fill="x", padx=28, pady=(14, 0))
        cheap, dear, budget = data["cheapest"], data["dearest"], data["budget"]
        tiles = []
        if cheap is not None:
            tiles.append(("En düşük teklif", f"{format_amount(cheap['effective'])} ₺", cheap["label"], OK))
        else:
            tiles.append(("En düşük teklif", "—", "Tutar girilmemiş", MUTED))
        if cheap is not None and dear is not None and dear is not cheap:
            diff = dear["effective"] - cheap["effective"]
            pct = diff / cheap["effective"] * 100 if cheap["effective"] else 0
            tiles.append(("Teklifler arası fark", f"{format_amount(diff)} ₺", f"En yüksek {dear['label']} · %{pct:.0f}", WARN))
        else:
            tiles.append(("Teklifler arası fark", "—", "Tek fiyat var", MUTED))
        if budget is not None:
            priced = [entry for entry in data["entries"] if entry["effective"] is not None]
            inside = [entry for entry in priced if entry["effective"] <= budget]
            color = OK if len(inside) == len(priced) and priced else (DANGER if not inside else WARN)
            note = f"Bütçe {format_amount(budget)} ₺"
            if cheap is not None:
                gap = budget - cheap["effective"]
                note += f" · en düşük {format_amount(abs(gap))} ₺ {'altında' if gap >= 0 else 'üstünde'}"
            tiles.append(("Bütçe uyumu", f"{len(inside)} / {len(priced)}", note, color))
        else:
            tiles.append(("Bütçe uyumu", "—", "Projede bütçe yok", MUTED))
        match = data["match"]
        tiles.append(("Ürün eşleşmesi", match["short"], "Aynı ürün mü kıyaslanıyor", self.MATCH_COLORS[match["level"]][1]))
        for index, (title, value, note, color) in enumerate(tiles):
            tile = tk.Frame(row, bg=CARD, highlightbackground=LINE, highlightthickness=1)
            tile.pack(side="left", fill="both", expand=True, padx=(0 if index == 0 else 12, 0))
            tk.Frame(tile, bg=color, height=4).pack(fill="x")
            inner = tk.Frame(tile, bg=CARD)
            inner.pack(fill="both", expand=True, padx=16, pady=12)
            tk.Label(inner, text=title, bg=CARD, fg=MUTED, font=(FONT, 11), anchor="w").pack(fill="x")
            tk.Label(inner, text=value, bg=CARD, fg=INK, font=(FONT, 20, "bold"), anchor="w").pack(fill="x", pady=(2, 0))
            tk.Label(inner, text=note, bg=CARD, fg=MUTED, font=(FONT, 11), anchor="w", justify="left", wraplength=220).pack(fill="x")

    def _match(self, data, width):
        match = data["match"]
        bg, fg = self.MATCH_COLORS[match["level"]]
        box = tk.Frame(self.body, bg=bg)
        box.pack(fill="x", padx=28, pady=(14, 0))
        tk.Label(box, text=f"Ürünler aynı mı?  {match['short']}", bg=bg, fg=fg, font=(FONT, 12, "bold"), anchor="w").pack(fill="x", padx=16, pady=(10, 0))
        tk.Label(box, text=match["text"], bg=bg, fg=INK, font=(FONT, 12), anchor="w", justify="left", wraplength=width - 100).pack(fill="x", padx=16, pady=(2, 10))

    def _score_cards(self, data, width):
        wrap = tk.Frame(self.body, bg=CANVAS)
        wrap.pack(fill="x", padx=28, pady=(10, 0))
        ranked = data["ranked"]
        per_row = min(len(ranked), 3) or 1
        card_wrap = max(220, (width - 56) // per_row - 50)
        for column in range(per_row):
            wrap.columnconfigure(column, weight=1, uniform="score")
        for index, entry in enumerate(ranked):
            top = index == 0
            card = tk.Frame(wrap, bg=CARD, highlightbackground=ACCENT if top else LINE, highlightthickness=2 if top else 1)
            card.grid(row=index // per_row, column=index % per_row, sticky="nsew", padx=(0 if index % per_row == 0 else 12, 0), pady=(0, 12))
            head = tk.Frame(card, bg=CARD)
            head.pack(fill="x", padx=16, pady=(14, 0))
            badge_bg = ACCENT if top else "#eef2f6"
            tk.Label(head, text=f"{index + 1}.", bg=badge_bg, fg="white" if top else INK, font=(FONT, 12, "bold"), padx=8, pady=2).pack(side="left")
            tk.Label(head, text=entry["label"], bg=CARD, fg=INK, font=(FONT, 15, "bold"), anchor="w").pack(side="left", padx=(8, 0))
            tk.Label(head, text=str(entry["score"]), bg=CARD, fg=ACCENT if top else INK, font=(FONT, 22, "bold")).pack(side="right")
            bar = tk.Frame(card, bg=LINE, height=8)
            bar.pack(fill="x", padx=16, pady=(8, 0))
            fill = tk.Frame(bar, bg=ACCENT if top else MUTED)
            fill.place(x=0, y=0, relheight=1, relwidth=max(0.02, entry["score"] / 100))
            amount = "Tutar yok" if entry["amount"] is None else f"{format_amount(entry['amount'])} ₺"
            if entry["profile"]["kdv"] == "haric" and entry["effective"] is not None:
                amount += f"  (KDV ile {format_amount(entry['effective'])} ₺)"
            tk.Label(card, text=amount, bg=CARD, fg=INK, font=(FONT, 14), anchor="w").pack(fill="x", padx=16, pady=(10, 0))
            parts = "   ".join(f"{COMPARE_LABELS[key]} {round(value * 100)}" for key, value in entry["parts"].items())
            if parts:
                tk.Label(card, text=parts, bg=CARD, fg=MUTED, font=(FONT, 11), anchor="w").pack(fill="x", padx=16, pady=(2, 0))
            tk.Label(card, text="Fiyat / fayda", bg=CARD, fg=ACCENT, font=(FONT, 11, "bold"), anchor="w").pack(fill="x", padx=16, pady=(10, 0))
            tk.Label(card, text=entry["value"], bg=CARD, fg=INK, font=(FONT, 12), anchor="w", justify="left", wraplength=card_wrap).pack(fill="x", padx=16, pady=(2, 0))
            actions = tk.Frame(card, bg=CARD)
            actions.pack(fill="x", padx=16, pady=(10, 14))
            if entry["row"]["file_path"]:
                Click(actions, "Teklif belgesini aç", lambda path=entry["row"]["file_path"]: self.app._open_offer_file(path), SOFT, INK, SOFT_HOVER, font=(FONT, 11), padx=10, pady=4).pack(side="left", padx=(0, 6))
            page = (entry["row"]["url"] if "url" in entry["row"].keys() else "") or ""
            if page:
                Click(actions, "Sayfayı aç", lambda url=page: self.app._open_offer_url(url), SOFT, INK, SOFT_HOVER, font=(FONT, 11), padx=10, pady=4).pack(side="left")

    def _matrix(self, data, width):
        entries = data["entries"]
        budget = data["budget"]
        table = tk.Frame(self.body, bg=LINE)
        table.pack(fill="x", padx=28, pady=(10, 0))
        table.columnconfigure(0, weight=0, minsize=190)
        for column in range(len(entries)):
            table.columnconfigure(column + 1, weight=1, uniform="offer")
        cell_wrap = max(140, (width - 56 - 190) // max(1, len(entries)) - 24)
        tk.Label(table, text="Kriter", bg=NAV, fg="white", font=(FONT, 12, "bold"), anchor="w", padx=12, pady=10).grid(row=0, column=0, sticky="nsew", padx=(1, 1), pady=(1, 1))
        pick = data["verdict"]["pick"]
        for column, entry in enumerate(entries):
            title = entry["label"] + ("  ★" if entry is pick else "")
            tk.Label(table, text=title, bg=ACCENT if entry is pick else NAV, fg="white", font=(FONT, 12, "bold"), anchor="w", padx=12, pady=10, wraplength=cell_wrap, justify="left").grid(row=0, column=column + 1, sticky="nsew", padx=(0, 1), pady=(1, 1))

        def money(entry):
            return f"{format_amount(entry['amount'])} ₺" if entry["amount"] is not None else "—"

        def compared(entry):
            if entry["effective"] is None:
                return "—"
            suffix = "  (KDV eklendi)" if entry["profile"]["kdv"] == "haric" else ""
            return f"{format_amount(entry['effective'])} ₺{suffix}"

        def budget_text(entry):
            if budget is None or entry["effective"] is None:
                return "—"
            gap = budget - entry["effective"]
            return f"{format_amount(gap)} ₺ altında" if gap >= 0 else f"{format_amount(-gap)} ₺ aşıyor"

        def kdv_text(entry):
            return {"dahil": "Dahil", "haric": "Hariç"}.get(entry["profile"]["kdv"], "Yazmıyor")

        def warranty_text(entry):
            return _months_text(entry["profile"]["warranty_months"]) or "Yazmıyor"

        def delivery_text(entry):
            if entry["profile"]["delivery_days"]:
                return f"{entry['profile']['delivery_days']} gün" + ("  · stokta" if entry["profile"]["stock"] else "")
            return "Stokta, hemen" if entry["profile"]["stock"] else "Yazmıyor"

        def payment_text(entry):
            days = entry["profile"]["payment_days"]
            if days is None:
                return "Yazmıyor"
            return "Peşin" if days == 0 else f"{days} gün vade"

        def labels(words):
            return ", ".join(label for word, label in OFFER_FEATURES if word in words) or "—"

        def validity_text(entry):
            days = entry["profile"]["validity_days"]
            return f"{days} gün" if days else "Yazmıyor"

        def discount_text(entry):
            return f"%{entry['profile']['discount']}" if entry["profile"]["discount"] else "—"

        def spec_text(entry):
            return ", ".join(entry["specs"]) or "Okunmadı"

        def document_text(entry):
            page = (entry["row"]["url"] if "url" in entry["row"].keys() else "") or ""
            bits = []
            if entry["row"]["file_path"]:
                kind = offer_kind(entry["row"]["file_path"])
                bits.append(kind if entry["profile"]["has_text"] else f"{kind} · metin okunamadı")
            if page:
                bits.append(f"Sayfa · {offer_url_host(page)}")
            return "\n".join(bits) or "Belge yok"

        def rank(values, better):
            present = [value for value in values if value is not None]
            if len(set(present)) < 2:
                return {}
            best = min(present) if better == "low" else max(present)
            worst = max(present) if better == "low" else min(present)
            marks = {}
            for index, value in enumerate(values):
                if value == best:
                    marks[index] = "best"
                elif value == worst or value is None:
                    marks[index] = "worst"
            return marks

        rows = [
            ("Teklif tutarı", money, None),
            ("Kıyas tutarı", compared, rank([entry["effective"] for entry in entries], "low")),
            ("Bütçeye göre", budget_text, rank([None if budget is None or entry["effective"] is None else budget - entry["effective"] for entry in entries], "high")),
            ("Teklif adı", lambda entry: entry["row"]["title"] or "—", None),
        ]
        hw_labels = [label for label in HW_ORDER if any(label in entry["hw"] for entry in entries)]
        for label in hw_labels:
            values = [entry["hw"].get(label, {}).get("value") for entry in entries]
            marks = rank(values, "high") if any(value is not None for value in values) else None
            rows.append((label, lambda entry, label=label: entry["hw"].get(label, {}).get("text") or "Yazmıyor", marks))
        if not hw_labels:
            rows.append(("Teknik özellikler", spec_text, None))
        for key, label in shared_spec_rows(entries):
            rows.append((label, lambda entry, key=key: entry["table"].get(key, ("", "Yazmıyor"))[1], None))
        rows += [
            ("Garanti", warranty_text, rank([entry["profile"]["warranty_months"] for entry in entries], "high")),
            ("Teslim süresi", delivery_text, rank([entry["delivery"] for entry in entries], "low")),
            ("Ödeme", payment_text, rank([entry["profile"]["payment_days"] for entry in entries], "high")),
            ("KDV", kdv_text, None),
            ("Fiyata dahil", lambda entry: labels(entry["profile"]["includes"]), rank([len(entry["profile"]["includes"]) for entry in entries], "high")),
            ("Fiyata dahil değil", lambda entry: labels(entry["profile"]["excludes"]), None),
            ("İskonto", discount_text, None),
            ("Teklif geçerliliği", validity_text, None),
            ("Puan", lambda entry: f"{entry['score']} / 100", rank([entry["score"] for entry in entries], "high")),
            ("Belge", document_text, None),
        ]
        for line, (label, render, marks) in enumerate(rows, start=1):
            stripe = CARD if line % 2 else "#f7f9fc"
            tk.Label(table, text=label, bg=stripe, fg=MUTED, font=(FONT, 12, "bold"), anchor="nw", padx=12, pady=8).grid(row=line, column=0, sticky="nsew", padx=(1, 1), pady=(0, 1))
            for column, entry in enumerate(entries):
                state = (marks or {}).get(column, "")
                bg = {"best": OK_SOFT, "worst": DANGER_SOFT}.get(state, stripe)
                fg = {"best": OK, "worst": DANGER}.get(state, INK)
                weight = "bold" if state == "best" else "normal"
                tk.Label(table, text=render(entry), bg=bg, fg=fg, font=(FONT, 12, weight), anchor="nw", justify="left", padx=12, pady=8, wraplength=cell_wrap).grid(row=line, column=column + 1, sticky="nsew", padx=(0, 1), pady=(0, 1))

    def _gaps(self, data, width):
        box = tk.Frame(self.body, bg=CARD, highlightbackground=LINE, highlightthickness=1)
        box.pack(fill="x", padx=28, pady=(10, 0))
        gaps = data["gaps"] or ["Teklifler aynı başlıklarda bilgi veriyor; eksik görünmüyor."]
        for item in gaps:
            line = tk.Frame(box, bg=CARD)
            line.pack(fill="x", padx=16, pady=(8, 0))
            tk.Label(line, text="●", bg=CARD, fg=WARN if data["gaps"] else OK, font=(FONT, 10)).pack(side="left", anchor="n", pady=(3, 0))
            tk.Label(line, text=item, bg=CARD, fg=INK, font=(FONT, 12), anchor="w", justify="left", wraplength=width - 120).pack(side="left", fill="x", padx=(8, 0))
        tk.Frame(box, bg=CARD, height=10).pack(fill="x")


class OfferDialog(tk.Toplevel):
    def __init__(self, app, project, row=None):
        super().__init__(app)
        self.app = app
        self.project = project
        self.row = row
        self.result = None
        self.title("Teklifi düzenle" if row else "Teklif ekle")
        self.configure(bg=CANVAS)
        self.transient(app)
        self.resizable(False, False)
        card = tk.Frame(self, bg=CARD, highlightbackground=LINE, highlightthickness=1)
        card.pack(fill="both", expand=True, padx=16, pady=16)
        tk.Label(card, text=project["name"], bg=CARD, fg=ACCENT, font=(FONT, 11, "bold")).pack(anchor="w", padx=16, pady=(14, 0))
        tk.Label(card, text="Teklif ekle" if row is None else "Teklifi düzenle", bg=CARD, fg=INK, font=(FONT, 18, "bold")).pack(anchor="w", padx=16)
        tk.Label(card, text="Dosya ya da internet sayfası eklenince özellikler ve koşullar okunur. Avantaj ve dezavantaj her firmaya göre ayrı yazılır.", bg=CARD, fg=MUTED, font=(FONT, 12), wraplength=560, justify="left").pack(anchor="w", padx=16, pady=(2, 8))
        form = tk.Frame(card, bg=CARD)
        form.pack(fill="x", padx=16)
        form.columnconfigure(1, weight=1)
        self.vendor = tk.StringVar(value=(row["vendor"] if row else ""))
        self.offer_title = tk.StringVar(value=(row["title"] if row else "") or "")
        amount = format_amount(row["amount"]) if row and row["amount"] is not None else ""
        self.amount = tk.StringVar(value=amount)
        self.offer_date = tk.StringVar(value=(row["offer_date"] if row else "") or "")
        self.file_path = tk.StringVar(value=(row["file_path"] if row else "") or "")
        self.page_url = tk.StringVar(value=((row["url"] if row and "url" in row.keys() else "") or ""))
        self.extracted = (row["source_text"] if row and "source_text" in row.keys() else "") or ""
        self.file_text = None
        self.url_text = None
        self._read_token = 0
        self._url_token = 0
        self._read_url = self.page_url.get().strip()
        self._url_busy = False
        self._save_waiting = False
        self._field(form, 0, "Firma", self.vendor)
        self._field(form, 1, "Teklif adı", self.offer_title)
        self._field(form, 2, "Tutar (₺)", self.amount)
        tk.Label(form, text="Tarih", bg=CARD, fg=MUTED, font=(FONT, 12)).grid(row=3, column=0, sticky="w", pady=4)
        date_row = tk.Frame(form, bg=CARD)
        date_row.grid(row=3, column=1, sticky="ew", pady=4)
        tk.Entry(date_row, textvariable=self.offer_date, width=18, relief="flat", highlightthickness=1, highlightbackground=LINE, font=(FONT, 12)).pack(side="left", ipady=3)
        Click(date_row, "Takvim", lambda: CalendarPopup(self, self.offer_date), SOFT, INK, SOFT_HOVER, font=(FONT, 12), padx=10, pady=4).pack(side="left", padx=(6, 0))
        tk.Label(form, text="Dosya", bg=CARD, fg=MUTED, font=(FONT, 12)).grid(row=4, column=0, sticky="nw", pady=4)
        file_row = tk.Frame(form, bg=CARD)
        file_row.grid(row=4, column=1, sticky="ew", pady=4)
        self.file_label = tk.Label(file_row, text=self._file_caption(), bg=CARD, fg=INK, font=(FONT, 12), anchor="w", wraplength=360, justify="left")
        self.file_label.pack(anchor="w")
        picks = tk.Frame(file_row, bg=CARD)
        picks.pack(anchor="w", pady=(4, 0))
        Click(picks, "Dosya seç", self._pick_file, SOFT, INK, SOFT_HOVER, font=(FONT, 12), padx=10, pady=4).pack(side="left")
        Click(picks, "Kaldır", self._clear_file, CARD, DANGER, "#fdecea", font=(FONT, 12), padx=10, pady=4, border=LINE).pack(side="left", padx=(6, 0))
        tk.Label(form, text="İnternet sayfası", bg=CARD, fg=MUTED, font=(FONT, 12)).grid(row=5, column=0, sticky="nw", pady=4)
        url_box = tk.Frame(form, bg=CARD)
        url_box.grid(row=5, column=1, sticky="ew", pady=4)
        url_line = tk.Frame(url_box, bg=CARD)
        url_line.pack(fill="x")
        self.url_entry = tk.Entry(url_line, textvariable=self.page_url, relief="flat", highlightthickness=1, highlightbackground=LINE, font=(FONT, 12))
        self.url_entry.pack(side="left", fill="x", expand=True, ipady=3)
        self.url_entry.bind("<Return>", lambda _e: self._read_page())
        self.url_entry.bind("<<Paste>>", lambda _e: self.after(50, self._read_page))
        self.url_entry.bind("<FocusOut>", lambda _e: self._read_page() if self._url_pending() else None)
        Click(url_line, "Oku", self._read_page, SOFT, INK, SOFT_HOVER, font=(FONT, 12), padx=10, pady=4).pack(side="left", padx=(6, 0))
        Click(url_line, "Aç", self._open_page, CARD, INK, HOVER, font=(FONT, 12), padx=10, pady=4, border=LINE).pack(side="left", padx=(6, 0))
        self.url_label = tk.Label(url_box, text=self._url_caption(), bg=CARD, fg=MUTED, font=(FONT, 11), anchor="w", wraplength=360, justify="left")
        self.url_label.pack(anchor="w", pady=(3, 0))
        tk.Label(form, text="Avantajlar", bg=CARD, fg=OK, font=(FONT, 12, "bold")).grid(row=6, column=0, sticky="nw", pady=4)
        self.pros = tk.Text(form, height=4, width=42, relief="flat", highlightthickness=1, highlightbackground=LINE, font=(FONT, 12), wrap="word")
        self.pros.grid(row=6, column=1, sticky="ew", pady=4)
        tk.Label(form, text="Dezavantajlar", bg=CARD, fg=DANGER, font=(FONT, 12, "bold")).grid(row=7, column=0, sticky="nw", pady=4)
        self.cons = tk.Text(form, height=4, width=42, relief="flat", highlightthickness=1, highlightbackground=LINE, font=(FONT, 12), wrap="word")
        self.cons.grid(row=7, column=1, sticky="ew", pady=4)
        if row:
            self.pros.insert("1.0", row["pros"] or "")
            self.cons.insert("1.0", row["cons"] or "")
        tk.Label(form, text="Not", bg=CARD, fg=MUTED, font=(FONT, 12)).grid(row=8, column=0, sticky="nw", pady=4)
        self.notes = tk.Text(form, height=2, width=42, relief="flat", highlightthickness=1, highlightbackground=LINE, font=(FONT, 12), wrap="word")
        self.notes.grid(row=8, column=1, sticky="ew", pady=4)
        if row and row["notes"]:
            self.notes.insert("1.0", row["notes"])
        buttons = tk.Frame(card, bg=CARD)
        buttons.pack(fill="x", padx=16, pady=14)
        Click(buttons, "Vazgeç", self.destroy, CANVAS, INK, HOVER, font=(FONT, 12), padx=12, pady=6, border=LINE).pack(side="right")
        Click(buttons, "Kaydet", self._save, ACCENT, "white", ACCENT_HOVER, font=(FONT, 12, "bold"), padx=14, pady=6).pack(side="right", padx=(0, 8))
        self.bind("<Escape>", lambda _e: self.destroy())
        self.update_idletasks()
        self.geometry(f"+{app.winfo_rootx() + 80}+{max(20, app.winfo_rooty() + 30)}")
        self.grab_set()
        if self._read_url and SPEC_HEADER not in self.extracted:
            self.after(200, self._read_page)

    def _url_pending(self):
        url = self.page_url.get().strip()
        if not url or self._url_busy:
            return False
        try:
            url = clean_offer_url(url)
        except ValueError:
            return False
        return self.url_text is None or url != self._read_url

    def _field(self, form, row, label, variable):
        tk.Label(form, text=label, bg=CARD, fg=MUTED, font=(FONT, 12)).grid(row=row, column=0, sticky="w", pady=4)
        tk.Entry(form, textvariable=variable, relief="flat", highlightthickness=1, highlightbackground=LINE, font=(FONT, 12)).grid(row=row, column=1, sticky="ew", pady=4, ipady=3)

    def _file_caption(self):
        path = self.file_path.get().strip()
        if not path:
            return "Dosya seçilmedi"
        return f"{offer_kind(path)}  ·  {Path(path).name}"

    def _pick_file(self):
        path = filedialog.askopenfilename(
            parent=self,
            title="Teklif dosyası",
            filetypes=[
                ("Teklif", "*.pdf *.eml *.png *.jpg *.jpeg *.webp *.tif *.tiff *.heic *.docx *.xlsx"),
                ("PDF", "*.pdf"),
                ("E-posta", "*.eml"),
                ("Görüntü", "*.png *.jpg *.jpeg *.webp *.tif *.tiff *.heic"),
                ("Word", "*.docx"),
                ("Excel", "*.xlsx"),
            ],
        )
        if not path:
            return
        try:
            assert_offer_file(path)
        except ValueError as exc:
            messagebox.showwarning(APP_NAME, str(exc), parent=self)
            return
        self.file_path.set(path)
        self.file_text = None
        self._read_token += 1
        token = self._read_token
        self.file_label.configure(text="Dosya okunuyor…")
        threading.Thread(target=self._read_file, args=(path, token), daemon=True).start()

    def _siblings(self):
        siblings = []
        for item in self.app.db.offers_for(self.project["id"]):
            if self.row and item["id"] == self.row["id"]:
                continue
            siblings.append({
                "vendor": item["vendor"],
                "amount": item["amount"],
                "file_path": item["file_path"] or (item["url"] if "url" in item.keys() else ""),
                "source_text": item["source_text"] if "source_text" in item.keys() else "",
            })
        return siblings

    def _source_text(self, final=False):
        if self.file_text is None and self.url_text is None:
            return self.extracted
        file_part = self.file_text
        path = self.file_path.get().strip()
        if file_part is None:
            file_part = ""
            if final and path:
                try:
                    file_part = read_offer_text(path)
                except (OSError, ValueError, subprocess.SubprocessError, zipfile.BadZipFile, ET.ParseError):
                    file_part = ""
        url_part = self.url_text
        url = self.page_url.get().strip()
        if url_part is None or url != self._read_url:
            url_part = f"Adres: {offer_url_slug_text(url)}" if url else ""
        return "\n".join(part for part in (file_part, url_part) if part)[:20000]

    def _refresh_preview(self):
        vendor = self.vendor.get().strip() or "Firma"
        try:
            amount = parse_amount(self.amount.get(), False)
        except ValueError:
            amount = None
        pros, cons = preview_offer_text(vendor, amount, self._source_text(), self._siblings())
        if self._box_is_generated(self.pros):
            self._set_box(self.pros, pros)
        if self._box_is_generated(self.cons):
            self._set_box(self.cons, cons)

    def _url_caption(self):
        url = self.page_url.get().strip()
        if not url:
            return "Ürün sayfasının adresini yapıştır. Sayfa okunur; site izin vermezse bilgiler adresten çıkarılır."
        host = offer_url_host(url)
        return f"{host}  ·  kayıtlı" if host else "Adres kayıtlı"

    def _read_page(self):
        raw = self.page_url.get().strip()
        if not raw:
            self.url_text = ""
            self._read_url = ""
            self.url_label.configure(text=self._url_caption(), fg=MUTED)
            self._refresh_preview()
            return
        try:
            url = clean_offer_url(raw)
        except ValueError as exc:
            self.url_label.configure(text=str(exc), fg=DANGER)
            return
        if url != raw:
            self.page_url.set(url)
        self._url_token += 1
        token = self._url_token
        self._url_busy = True
        self.url_label.configure(text=f"{offer_url_host(url)} okunuyor… Teknik özellikler aranıyor.", fg=MUTED)
        threading.Thread(target=self._fetch_page, args=(url, token), daemon=True).start()

    def _fetch_page(self, url, token):
        try:
            page = fetch_offer_page(url)
        except ValueError as exc:
            page = {"url": url, "title": "", "price": None, "text": "", "blocked": True, "note": str(exc), "specs": []}

        def apply():
            if not self.winfo_exists() or token != self._url_token:
                return
            self._url_busy = False
            self.url_text = page["text"]
            self._read_url = url
            if page["price"] is not None and not self.amount.get().strip():
                self.amount.set(format_amount(page["price"]) or "")
            if page["title"] and not self.offer_title.get().strip():
                self.offer_title.set(page["title"][:160])
            found = offer_hardware(page["text"])
            parts = [f"{label}: {found[label]['short']}" for label in ("İşlemci", "Ekran kartı", "RAM", "Depolama") if label in found]
            caption = f"{offer_url_host(url)}  ·  {page['note']}"
            if parts:
                caption += "\nOkunan: " + ", ".join(parts)
            specs = page.get("specs") or []
            if specs:
                caption += f"\n{len(specs)} teknik özellik bulundu ve teklife eklendi."
            elif not page["blocked"]:
                caption += "\nSayfada teknik özellik tablosu bulunamadı; başlık ve açıklama kullanıldı."
            if page["price"] is not None:
                caption += f"\nSayfadaki fiyat: {format_amount(page['price'])} ₺"
            self.url_label.configure(text=caption, fg=WARN if page["blocked"] else OK)
            self._refresh_preview()
            if self._save_waiting:
                self._save_waiting = False
                self._save()

        self.after(0, apply)

    def _open_page(self):
        try:
            url = clean_offer_url(self.page_url.get())
        except ValueError as exc:
            self.url_label.configure(text=str(exc), fg=DANGER)
            return
        if url:
            webbrowser.open(url)

    def _box_is_generated(self, widget):
        text = widget.get("1.0", "end").strip()
        if not text:
            return True
        lines = [line.strip() for line in text.splitlines() if line.strip()]
        return bool(lines) and all(line.startswith("•") for line in lines)

    def _set_box(self, widget, text):
        widget.delete("1.0", "end")
        if text:
            widget.insert("1.0", text)

    def _read_file(self, path, token):
        try:
            text = read_offer_text(path)
        except (OSError, ValueError, subprocess.SubprocessError, zipfile.BadZipFile, ET.ParseError):
            text = ""
        total = find_offer_total(fold(text))

        def apply():
            if not self.winfo_exists() or token != self._read_token:
                return
            self.file_text = text
            if total is not None and not self.amount.get().strip():
                self.amount.set(format_amount(total) or "")
            self._refresh_preview()
            self.file_label.configure(text=self._file_caption() + "  ·  koşullar yazıldı")

        self.after(0, apply)

    def _clear_file(self):
        self._read_token += 1
        self.file_path.set("")
        self.file_text = ""
        self.file_label.configure(text=self._file_caption())

    def _save(self):
        vendor = self.vendor.get().strip()
        if not vendor:
            messagebox.showwarning(APP_NAME, "Firma boş bırakılamaz.", parent=self)
            return
        if len(vendor) > 160:
            messagebox.showwarning(APP_NAME, "Firma adı çok uzun.", parent=self)
            return
        title = self.offer_title.get().strip()
        if len(title) > 160:
            messagebox.showwarning(APP_NAME, "Teklif adı çok uzun.", parent=self)
            return
        day = self.offer_date.get().strip()
        if day and parse_day(day) is None:
            messagebox.showwarning(APP_NAME, "Tarih için YYYY-AA-GG kullan.", parent=self)
            return
        try:
            amount = parse_amount(self.amount.get(), False)
        except ValueError as exc:
            messagebox.showwarning(APP_NAME, str(exc), parent=self)
            return
        pros = self.pros.get("1.0", "end").strip()
        cons = self.cons.get("1.0", "end").strip()
        notes = self.notes.get("1.0", "end").strip()
        if len(pros) > 2000 or len(cons) > 2000 or len(notes) > 1000:
            messagebox.showwarning(APP_NAME, "Metin çok uzun.", parent=self)
            return
        try:
            url = clean_offer_url(self.page_url.get())
        except ValueError as exc:
            messagebox.showwarning(APP_NAME, str(exc), parent=self)
            return
        if url and (self._url_busy or self._url_pending()):
            self._save_waiting = True
            if not self._url_busy:
                self._read_page()
            self.url_label.configure(text=f"{offer_url_host(url)} okunuyor… Bitince teklif kaydedilecek.", fg=MUTED)
            return
        previous_url = ((self.row["url"] if self.row and "url" in self.row.keys() else "") or "")
        stored = ""
        chosen = self.file_path.get().strip()
        if chosen:
            try:
                stored = store_offer_file(chosen)
            except ValueError as exc:
                messagebox.showwarning(APP_NAME, str(exc), parent=self)
                return
        self.result = {
            "project_id": int(self.project["id"]),
            "vendor": vendor,
            "title": title,
            "amount": amount,
            "offer_date": day,
            "file_path": stored,
            "url": url,
            "pros": pros,
            "cons": cons,
            "notes": notes,
        }
        if url != previous_url and self.url_text is None:
            self.url_text = None
            self._read_url = ""
            if self.file_text is None and not chosen:
                self.file_text = ""
        changed = self.file_text is not None or self.url_text is not None or url != previous_url
        if changed:
            if self.file_text is None and self.url_text is None:
                self.url_text = ""
            self.result["source_text"] = self._source_text(final=True)
        elif self.extracted:
            self.result["source_text"] = self.extracted[:20000]
        self.destroy()


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        apply_window_icon(self)
        set_dock_icon()
        threading.Thread(target=self._warm_report_tool, daemon=True).start()
        self.configure(bg=CANVAS)
        self.minsize(1100, 700)
        self._alive = True
        self._locked_out = False
        self._mail_busy = False
        self.events = queue.Queue()
        self.db = Database()
        self._lock_file = open(DATA_DIR / ".lock", "a+b")
        if not try_single_instance(self._lock_file):
            self.db.close()
            messagebox.showinfo(APP_NAME, "VKV İş Takip Merkezi zaten açık.")
            self._locked_out = True
            self.after(50, self.destroy)
            return
        self.section = "İş takibi"
        self.board_project_id = None
        self.db.ensure_workspace()
        self._column_filters = {}
        self.search_var = tk.StringVar()
        self.invoice_filter = tk.StringVar(value="Tümü")
        self.invoice_sort = tk.StringVar(value="Tarih (yeni)")
        self.selected_invoice = None
        self.nav_buttons = {}
        self.metric_values = {}
        self.metric_notes = {}
        self._apply_style()
        self._build()
        self.search_var.trace_add("write", self._on_search_changed)
        self._place()
        try:
            note = self.sync_folder()
        except (OSError, ValueError, zipfile.BadZipFile) as exc:
            note = f"Klasör henüz okunamadı: {exc}"
        self.show_section("İş takibi")
        if note:
            self.status.configure(text=note)
        self.after(800, self.check_due)
        self.after(1500, self.tick_folder)
        self.after(200, self._drain)
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.after(150, self._bring_front)

    def _apply_style(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure(".", font=(FONT, 13), background=CANVAS)
        style.configure("TFrame", background=CANVAS)
        style.configure("Treeview", rowheight=42, font=(FONT, 13), background=CARD, fieldbackground=CARD, borderwidth=0, relief="flat")
        style.configure("Treeview.Item", padding=(4, 0))
        style.configure("Treeview.Heading", font=(FONT, 12), background=PAPER, foreground=MUTED, relief="flat", padding=(4, 8), anchor="w")
        style.layout("Treeview.Heading", [
            ("Treeheading.cell", {"sticky": "nswe"}),
            ("Treeheading.border", {"sticky": "nswe", "children": [
                ("Treeheading.padding", {"sticky": "nswe", "children": [
                    ("Treeheading.text", {"sticky": "we"}),
                ]}),
            ]}),
        ])
        style.layout("Treeview.Item", [
            ("Treeitem.padding", {"sticky": "nswe", "children": [
                ("Treeitem.text", {"sticky": "w"}),
            ]}),
        ])
        style.map("Treeview", background=[("selected", SOFT)], foreground=[("selected", INK)])
        style.map("Treeview.Heading", background=[("active", PAPER)], foreground=[("active", INK)])
        style.configure("TCombobox", padding=4)
        self.tree_tags_ready = False

    def _build(self):
        self.sidebar = tk.Frame(self, bg=NAV, width=236)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)
        brand = tk.Frame(self.sidebar, bg=NAV)
        brand.pack(fill="x", padx=18, pady=(22, 20))
        mark = tk.Canvas(brand, width=36, height=36, bg=NAV, highlightthickness=0)
        mark.pack(side="left")
        _round_rect(mark, 0, 0, 36, 36, 10, fill=ACCENT, outline=ACCENT)
        mark.create_line(11, 18, 16, 23, 26, 13, fill="white", width=2, capstyle="round", joinstyle="round")
        names = tk.Frame(brand, bg=NAV)
        names.pack(side="left", padx=(12, 0))
        tk.Label(names, text="VKV", bg=NAV, fg="white", font=(FONT, 15, "bold")).pack(anchor="w")
        tk.Label(names, text="İş Takip", bg=NAV, fg=NAV_MUTED, font=(FONT, 11)).pack(anchor="w")
        for name in PAGES:
            button = NavItem(self.sidebar, PAGE_TITLES.get(name, name), lambda page=name: self.show_section(page))
            button.pack(fill="x", padx=8, pady=1)
            self.nav_buttons[name] = button
            if name == "Faturalar":
                reports = NavItem(self.sidebar, "Raporlar", lambda: self.show_section("Raporlar"))
                reports.pack(fill="x", padx=8, pady=1)
                self.nav_buttons["Raporlar"] = reports
        Click(self.sidebar, "Ayarlar", self.open_settings, NAV, NAV_MUTED, NAV_HOVER, font=(FONT, 13), padx=14, pady=9, anchor="w").pack(side="bottom", fill="x", padx=22, pady=(0, 16))
        self.mail_state = tk.Label(self.sidebar, text="E-posta kapalı", bg=NAV, fg=NAV_MUTED, font=(FONT, 11), wraplength=180, justify="left")
        self.mail_state.pack(side="bottom", anchor="w", padx=26, pady=(0, 8))
        self._refresh_mail_label()

        self.main = tk.Frame(self, bg=CANVAS)
        self.main.pack(side="left", fill="both", expand=True)
        header = tk.Frame(self.main, bg=CANVAS)
        header.pack(fill="x", padx=24, pady=(20, 10))
        titles = tk.Frame(header, bg=CANVAS)
        titles.pack(side="left")
        self.heading = tk.Label(titles, text="İş takibi", bg=CANVAS, fg=INK, font=(FONT, 26, "bold"))
        self.heading.pack(anchor="w")
        self.subhead = tk.Label(titles, text=long_today(), bg=CANVAS, fg=MUTED, font=(FONT, 13))
        self.subhead.pack(anchor="w")
        self.search_box = tk.Frame(header, bg=CARD, highlightbackground=LINE, highlightthickness=1, cursor="xterm")
        search_label = tk.Label(self.search_box, text="Ara", bg=CARD, fg=MUTED, font=(FONT, 12), cursor="xterm")
        search_label.pack(side="left", padx=(12, 4))
        self.search = tk.Entry(self.search_box, textvariable=self.search_var, width=22, relief="flat", bg=CARD, fg=INK, highlightthickness=0, borderwidth=0, font=(FONT, 13), insertbackground=INK)
        self.search.pack(side="left", ipady=8, padx=(0, 12))
        self.search_box.pack(side="right")
        for widget in (self.search_box, search_label):
            widget.bind("<Button-1>", lambda _e: self.search.focus_set())
        self.search.bind("<KeyRelease>", self._on_search_changed)

        self.body = tk.Frame(self.main, bg=CANVAS)
        self.body.pack(fill="both", expand=True)
        self._build_board()
        self._build_invoices()
        self._build_reports()
        self._build_table()
        self.status = tk.Label(self.main, text="", bg=CANVAS, fg=MUTED, anchor="w", font=(FONT, 11))
        self.status.pack(fill="x", padx=22, pady=(4, 10))
        self.bind("<Command-n>", lambda _e: self.add_record())
        self.bind("<Command-f>", lambda _e: self.search.focus_set())

    def _build_board(self):
        self.board = tk.Frame(self.body, bg=CANVAS)
        metrics = tk.Frame(self.board, bg=CANVAS)
        metrics.pack(fill="x", padx=22, pady=(2, 14))
        cards = (
            ("open", "Açık iş", SOFT, ACCENT, "▣"),
            ("late", "Geciken", DANGER_SOFT, DANGER, "!"),
            ("unpaid", "Ödenmemiş", WARN_SOFT, WARN, "₺"),
            ("net", "Bu ay net", OK_SOFT, OK, "↗"),
        )
        for index, (key, title, tint, color, glyph) in enumerate(cards):
            card = tk.Frame(metrics, bg=CARD, highlightbackground=LINE, highlightthickness=1)
            card.pack(side="left", fill="both", expand=True, padx=(0, 0 if index == 3 else 12))
            inner = tk.Frame(card, bg=CARD)
            inner.pack(fill="both", expand=True, padx=16, pady=14)
            top = tk.Frame(inner, bg=CARD)
            top.pack(fill="x")
            badge = tk.Canvas(top, width=36, height=36, bg=CARD, highlightthickness=0)
            badge.pack(side="left")
            badge.create_oval(1, 1, 35, 35, fill=tint, outline="")
            badge.create_text(18, 18, text=glyph, fill=color, font=(FONT, 13, "bold"))
            tk.Label(top, text=title, bg=CARD, fg=MUTED, font=(FONT, 12)).pack(side="left", padx=(10, 0))
            value = tk.Label(inner, text="—", bg=CARD, fg=INK, font=(FONT, 26, "bold"))
            value.pack(anchor="w", pady=(10, 0))
            note = tk.Label(inner, text="", bg=CARD, fg=MUTED, font=(FONT, 11))
            note.pack(anchor="w", pady=(2, 0))
            self.metric_values[key] = value
            self.metric_notes[key] = note
        middle = tk.Frame(self.board, bg=CANVAS)
        middle.pack(fill="both", expand=True, padx=22, pady=(0, 12))
        self.columns = {}
        self.column_counts = {}
        board_wrap = tk.Frame(middle, bg=CANVAS)
        board_wrap.pack(side="left", fill="both", expand=True, padx=(0, 12))
        column_colors = {"Bekliyor": (WARN_SOFT, WARN), "Devam ediyor": (SOFT, ACCENT), "Tamamlandı": (OK_SOFT, OK)}
        for index, name in enumerate(("Bekliyor", "Devam ediyor", "Tamamlandı")):
            tint, color = column_colors[name]
            column = tk.Frame(board_wrap, bg=CARD, highlightbackground=LINE, highlightthickness=1)
            column.pack(side="left", fill="both", expand=True, padx=(0, 0 if index == 2 else 10))
            head = tk.Frame(column, bg=CARD)
            head.pack(fill="x", padx=12, pady=(12, 8))
            dot = tk.Canvas(head, width=10, height=10, bg=CARD, highlightthickness=0)
            dot.pack(side="left", pady=4)
            dot.create_oval(1, 1, 9, 9, fill=color, outline="")
            tk.Label(head, text=name, bg=CARD, fg=INK, font=(FONT, 13, "bold")).pack(side="left", padx=(8, 0))
            count = tk.Label(head, text="0", bg=tint, fg=color, font=(FONT, 11, "bold"), padx=8, pady=1)
            count.pack(side="right")
            self.column_counts[name] = count
            if name == "Bekliyor":
                Click(head, "+ Yeni", self.add_record, ACCENT, "white", ACCENT_HOVER, font=(FONT, 11, "bold"), padx=10, pady=3).pack(side="right", padx=(0, 8))
            holder = tk.Frame(column, bg=tint)
            holder.pack(fill="both", expand=True, padx=12, pady=(0, 10))
            self.columns[name] = holder
        self.board_preview = InvoicePreview(middle, self)
        self.board_preview.pack(side="right", fill="y")
        recent_wrap = tk.Frame(self.board_preview, bg=CARD)
        recent_wrap.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        self.recent_canvas = tk.Canvas(recent_wrap, bg=CARD, highlightthickness=0, height=180)
        self.recent_scroll = ttk.Scrollbar(recent_wrap, orient="vertical", command=self.recent_canvas.yview)
        self.recent_canvas.configure(yscrollcommand=self.recent_scroll.set)
        self.recent_canvas.pack(side="left", fill="both", expand=True)
        self.recent_scroll.pack(side="right", fill="y")
        self.recent = tk.Frame(self.recent_canvas, bg=CARD)
        self.recent_window = self.recent_canvas.create_window((0, 0), window=self.recent, anchor="nw")
        self.recent.bind("<Configure>", lambda _e: self.recent_canvas.configure(scrollregion=self.recent_canvas.bbox("all")))
        self.recent_canvas.bind("<Configure>", lambda event: self.recent_canvas.itemconfigure(self.recent_window, width=event.width))
        self.recent_canvas.bind("<MouseWheel>", self._wheel_recent)
        self.recent.bind("<MouseWheel>", self._wheel_recent)


    def _build_invoices(self):
        self.invoice_page = tk.Frame(self.body, bg=CANVAS)
        bar = tk.Frame(self.invoice_page, bg=CANVAS)
        bar.pack(fill="x", padx=22, pady=(0, 8))
        self._button(bar, "Yeni", self.add_record).pack(side="left")
        self._button(bar, "Düzenle", self.edit_record, False).pack(side="left", padx=6)
        self._button(bar, "Sil", self.delete_record, False).pack(side="left")
        self._button(bar, "Klasörü tara", self.sync_now, False).pack(side="left", padx=6)
        self.mail_button = self._button(bar, "E-postayı al", self.sync_mail_now, False)
        self.mail_button.pack(side="left")
        self._button(bar, "Excel", self.export_excel, False).pack(side="right")
        self._button(bar, "E-posta ayarı", self.open_settings, False).pack(side="right", padx=(0, 6))
        controls = tk.Frame(self.invoice_page, bg=CANVAS)
        controls.pack(fill="x", padx=22, pady=(0, 10))
        self.filter_toggle = Click(controls, "Filtreleme", lambda: self.open_drop(self.filter_toggle, ("Tümü", "Ödenmedi", "Ödendi", "Klasör", "E-posta"), self.invoice_filter, self.choose_filter), CARD, INK, HOVER, font=(FONT, 12), padx=12, pady=7, border=LINE)
        self.filter_toggle.pack(side="left")
        self.sort_toggle = Click(controls, "Sıralama", lambda: self.open_drop(self.sort_toggle, ("Tarih (yeni)", "Tarih (eski)", "Tutar (yüksek)", "Tutar (düşük)", "Firma"), self.invoice_sort, self.choose_sort), CARD, INK, HOVER, font=(FONT, 12), padx=12, pady=7, border=LINE)
        self.sort_toggle.pack(side="left", padx=(8, 0))
        self._refresh_tool_labels()
        split = tk.Frame(self.invoice_page, bg=CANVAS)
        split.pack(fill="both", expand=True, padx=22, pady=(0, 8))
        list_wrap = tk.Frame(split, bg=CANVAS)
        list_wrap.pack(side="left", fill="both", expand=True, padx=(0, 10))
        self.invoice_canvas = tk.Canvas(list_wrap, bg=CANVAS, highlightthickness=0)
        self.invoice_list_scroll = ttk.Scrollbar(list_wrap, orient="vertical", command=self.invoice_canvas.yview)
        self.invoice_canvas.configure(yscrollcommand=self.invoice_list_scroll.set)
        self.invoice_canvas.pack(side="left", fill="both", expand=True)
        self.invoice_list_scroll.pack(side="right", fill="y")
        self.invoice_list = tk.Frame(self.invoice_canvas, bg=CANVAS)
        self.invoice_list_window = self.invoice_canvas.create_window((0, 0), window=self.invoice_list, anchor="nw")
        self.invoice_list.bind("<Configure>", lambda _e: self.invoice_canvas.configure(scrollregion=self.invoice_canvas.bbox("all")))
        self.invoice_canvas.bind("<Configure>", lambda event: self.invoice_canvas.itemconfigure(self.invoice_list_window, width=event.width))
        self.invoice_canvas.bind("<MouseWheel>", self._wheel_invoices)
        self.invoice_list.bind("<MouseWheel>", self._wheel_invoices)
        self.invoice_preview = InvoicePreview(split, self, detail=True)
        self.invoice_preview.pack(side="right", fill="y")

    def _build_reports(self):
        self.report_page = tk.Frame(self.body, bg=CANVAS)
        bar = tk.Frame(self.report_page, bg=CANVAS)
        bar.pack(fill="x", padx=22, pady=(0, 8))
        self._button(bar, "Rapor Hazırla", self.prepare_report).pack(side="left")
        self.report_list = tk.Frame(self.report_page, bg=CANVAS)
        self.report_list.pack(fill="both", expand=True, padx=22, pady=(0, 8))

    def _build_table(self):
        self.list_wrap = tk.Frame(self.body, bg=CANVAS)
        self.toolbar = tk.Frame(self.list_wrap, bg=CANVAS)
        self.toolbar.pack(fill="x", padx=22, pady=(0, 8))
        self._button(self.toolbar, "Yeni", self.add_record).pack(side="left")
        self._button(self.toolbar, "Düzenle", self.edit_record, False).pack(side="left", padx=6)
        self._button(self.toolbar, "Sil", self.delete_record, False).pack(side="left")
        self.offer_button = self._button(self.toolbar, "Teklif ekle", self.add_offer)
        self.offer_panel = tk.Frame(self.list_wrap, bg=CANVAS)
        self.contact_find = tk.Frame(self.list_wrap, bg=CANVAS)
        find_box = tk.Frame(self.contact_find, bg=CARD, highlightbackground=LINE, highlightthickness=1)
        find_box.pack(fill="x")
        tk.Label(find_box, text="Outlook adres defterinde ara", bg=CARD, fg=MUTED, font=(FONT, 12)).pack(side="left", padx=(12, 6))
        self.contact_query = tk.Entry(find_box, textvariable=self.search_var, relief="flat", bg=CARD, fg=INK, highlightthickness=0, font=(FONT, 13), insertbackground=INK)
        self.contact_query.pack(side="left", fill="x", expand=True, ipady=8, padx=(0, 12))
        find_box.bind("<Button-1>", lambda _e: self.contact_query.focus_set())
        self._button(self.toolbar, "Excel", self.export_excel, False).pack(side="right")
        self._button(self.toolbar, "Yedekle", self.backup, False).pack(side="right", padx=(0, 6))
        table = tk.Frame(self.list_wrap, bg=CARD, highlightbackground=LINE, highlightthickness=1)
        self.table_frame = table
        table.pack(fill="both", expand=True, padx=22, pady=(4, 0))
        self.tree = ttk.Treeview(table, show="headings", selectmode="browse")
        self.tree_scroll = ttk.Scrollbar(table, orient="vertical", command=self.tree.yview)
        self.tree_xscroll = ttk.Scrollbar(table, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=self._on_tree_yview, xscrollcommand=self.tree_xscroll.set)
        table.rowconfigure(0, weight=1)
        table.columnconfigure(0, weight=1)
        self.tree.grid(row=0, column=0, sticky="nsew", padx=(8, 0), pady=(8, 0))
        self.tree_scroll.grid(row=0, column=1, sticky="ns", pady=(8, 0), padx=(0, 4))
        self.tree_xscroll.grid(row=1, column=0, sticky="ew", padx=(8, 0), pady=(0, 8))
        self._cell_labels = {}
        self._paint_text = {}
        self.tree.tag_configure("even", background=CARD)
        self.tree.tag_configure("odd", background=PAPER)
        self.tree.tag_configure("done", background=OK_SOFT)
        self.tree.tag_configure("run", background=SOFT)
        self.tree.tag_configure("wait", background=WARN_SOFT)
        self.tree.tag_configure("late", background=DANGER_SOFT)
        self.tree.tag_configure("plain", background=CARD, foreground=INK)
        self.tree.tag_configure("book", background=SOFT, foreground=INK)
        self.tree.bind("<Double-1>", self._on_tree_double)
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)
        self.tree.bind("<Configure>", lambda _e: self._schedule_cell_paint())
        self.tree.bind("<ButtonRelease-1>", lambda _e: self._schedule_cell_paint(), add="+")
        self.tree.bind("<Delete>", lambda _e: self.delete_record())
        self.tree.bind("<Button-1>", self._on_tree_click, add="+")
        self.tree.bind("<Motion>", self._on_tree_motion)
        self.empty_wrap = tk.Frame(table, bg=CARD)
        self.empty_title = tk.Label(self.empty_wrap, text="", bg=CARD, fg=INK, font=(FONT, 18, "bold"))
        self.empty_title.pack()
        self.empty_note = tk.Label(self.empty_wrap, text="", bg=CARD, fg=MUTED, font=(FONT, 13))
        self.empty_note.pack(pady=(4, 12))
        Click(self.empty_wrap, "Yeni kayıt", self.add_record, ACCENT, "white", ACCENT_HOVER, font=(FONT, 12, "bold"), padx=16, pady=8).pack()
        self.detail = tk.Label(self.list_wrap, text="", bg=CANVAS, fg=MUTED, anchor="w", justify="left", font=(FONT, 12), wraplength=860)
        self.detail.pack(fill="x", padx=22, pady=(8, 0))

    def _button(self, parent, text, command, primary=True):
        if primary:
            return Click(parent, text, command, ACCENT, "white", ACCENT_HOVER, font=(FONT, 12, "bold"), padx=14, pady=7)
        return Click(parent, text, command, CARD, INK, HOVER, font=(FONT, 12), padx=12, pady=7, border=LINE)

    def _refresh_tool_labels(self):
        self.filter_toggle.set_text(f"Filtreleme   {self.invoice_filter.get()}   ▾")
        self.sort_toggle.set_text(f"Sıralama   {self.invoice_sort.get()}   ▾")

    def open_drop(self, anchor, options, variable, command):
        previous = getattr(self, "_drop", None)
        if previous is not None and previous.winfo_exists():
            previous.destroy()
        menu = tk.Toplevel(self)
        menu.overrideredirect(True)
        menu.configure(bg=CARD)
        try:
            menu.attributes("-topmost", True)
        except tk.TclError:
            pass
        shell = tk.Frame(menu, bg=CARD, highlightbackground=LINE, highlightthickness=1)
        shell.pack(fill="both", expand=True)
        tk.Frame(shell, bg=ACCENT, height=2).pack(fill="x")
        current = variable.get()
        for name in options:
            selected = name == current
            row = tk.Frame(shell, bg=SOFT if selected else CARD, cursor="hand2")
            row.pack(fill="x")
            label = tk.Label(row, text=name, bg=row["bg"], fg=INK, anchor="w", padx=16, pady=9, font=(FONT, 13, "bold" if selected else "normal"))
            label.pack(side="left", fill="x", expand=True)
            mark = tk.Label(row, text="✓" if selected else "", bg=row["bg"], fg=ACCENT, padx=14, font=(FONT, 13, "bold"))
            mark.pack(side="right")

            def paint(background, widgets=(row, label, mark)):
                for widget in widgets:
                    widget.configure(bg=background)

            def pick(_event=None, value=name, paint=paint):
                command(value)
                if menu.winfo_exists():
                    menu.destroy()

            for widget in (row, label, mark):
                widget.bind("<Enter>", lambda _e, paint=paint: paint(HOVER))
                widget.bind("<Leave>", lambda _e, paint=paint, background=SOFT if selected else CARD: paint(background))
                widget.bind("<Button-1>", pick)
        menu.update_idletasks()
        width = max(220, anchor.winfo_width())
        x = anchor.winfo_rootx()
        y = anchor.winfo_rooty() + anchor.winfo_height() + 6
        menu.geometry(f"{width}x{menu.winfo_reqheight()}+{x}+{y}")
        self._drop = menu
        menu.bind("<Escape>", lambda _e: menu.destroy())
        menu.focus_set()

    def _cell_display(self, row, key):
        if key == "issue_key":
            return self.issue_key(row)
        if key == "project_name":
            return self._project_name(row)
        if key == "_note":
            raw = row["note_json"] if "note_json" in row.keys() else ""
            return "●" if note_has_items(raw) else "+"
        if key == "_docs":
            raw = row["docs_json"] if "docs_json" in row.keys() else ""
            count = len(ledger_documents(raw))
            return f"{count} belge" if count else "—"
        if key == "offer_count":
            count = self.db.offer_count(row["id"])
            return f"{count} teklif" if count else "—"
        if key in MONEY_KEYS:
            return format_amount(row[key])
        if key not in row.keys() or row[key] is None or str(row[key]).strip() == "":
            return "(boş)"
        return str(row[key])

    def _active_column_filters(self):
        return getattr(self, "_column_filters", {}).get(self.section, {})

    def open_column_filter(self, key, label):
        previous = getattr(self, "_drop", None)
        if previous is not None and previous.winfo_exists():
            previous.destroy()
        saved = dict(self._active_column_filters())
        current = saved.pop(key, None)
        self._column_filters[self.section] = saved
        try:
            pool = self.visible_rows()
        finally:
            if current is not None:
                saved[key] = current
            self._column_filters[self.section] = saved
        counts = {}
        for row in pool:
            text = self._cell_display(row, key)
            counts[text] = counts.get(text, 0) + 1
        preferred = []
        for _title, field, kind, extra in self.schema()["fields"]:
            if field == key and kind == "choice":
                preferred = list(extra)
        options = [name for name in preferred if name in counts]
        for name in sorted(counts, key=lambda item: (item == "(boş)", fold(item))):
            if name not in options:
                options.append(name)
        active = next(iter(current)) if current and len(current) == 1 else "Tümü"
        menu = tk.Toplevel(self)
        menu.overrideredirect(True)
        menu.configure(bg=CARD)
        try:
            menu.attributes("-topmost", True)
        except tk.TclError:
            pass
        shell = tk.Frame(menu, bg=CARD, highlightbackground=LINE, highlightthickness=1)
        shell.pack(fill="both", expand=True)
        tk.Frame(shell, bg=ACCENT, height=2).pack(fill="x")

        def choose(value):
            filters = self._column_filters.setdefault(self.section, {})
            if value == "Tümü":
                filters.pop(key, None)
            else:
                filters[key] = {value}
            if menu.winfo_exists():
                menu.destroy()
            self.after(1, self.refresh_list)

        for name in ("Tümü", *options):
            selected = name == active
            row = tk.Frame(shell, bg=SOFT if selected else CARD, cursor="hand2")
            row.pack(fill="x")
            shown = name if name == "Tümü" else f"{name}    {counts.get(name, 0)}"
            text = tk.Label(row, text=shown, bg=row["bg"], fg=INK, anchor="w", padx=16, pady=9, font=(FONT, 13, "bold" if selected else "normal"))
            text.pack(side="left", fill="x", expand=True)
            mark = tk.Label(row, text="✓" if selected else "", bg=row["bg"], fg=ACCENT, padx=14, font=(FONT, 13, "bold"))
            mark.pack(side="right")

            def paint(background, widgets=(row, text, mark)):
                for widget in widgets:
                    widget.configure(bg=background)

            def pick(_event=None, value=name):
                choose(value)

            for widget in (row, text, mark):
                widget.bind("<Enter>", lambda _e, paint=paint: paint(HOVER))
                widget.bind("<Leave>", lambda _e, paint=paint, background=SOFT if selected else CARD: paint(background))
                widget.bind("<Button-1>", pick)
        menu.update_idletasks()
        x = self.tree.winfo_rootx()
        for column in self.tree["columns"]:
            if column == key:
                break
            x += int(self.tree.column(column, "width"))
        y = self.tree.winfo_rooty() + 34
        width = max(240, menu.winfo_reqwidth())
        x = min(max(8, x), self.winfo_screenwidth() - width - 12)
        menu.geometry(f"{width}x{menu.winfo_reqheight()}+{x}+{y}")
        self._drop = menu
        menu.bind("<Escape>", lambda _e: menu.destroy())
        menu.focus_set()

    def choose_filter(self, name):
        self.invoice_filter.set(name)
        self._refresh_tool_labels()
        self.refresh_invoices()

    def choose_sort(self, name):
        self.invoice_sort.set(name)
        self._refresh_tool_labels()
        self.refresh_invoices()

    def _place(self):
        self.update_idletasks()
        width, height = 1280, 820
        x = max(0, (self.winfo_screenwidth() - width) // 2)
        y = max(0, (self.winfo_screenheight() - height) // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")

    def _bring_front(self):
        self.lift()
        self.attributes("-topmost", True)
        self.after(250, lambda: self.attributes("-topmost", False))

    def allowed_roots(self):
        return [self.watch_dir(), ATTACH_DIR]

    def watch_dir(self):
        raw = self.db.get_setting("watch_dir", "")
        path = Path(raw).expanduser() if raw else default_watch_dir()
        return path.resolve() if path.exists() else path

    def excel_path(self):
        raw = self.db.get_setting("excel_path", "")
        if raw:
            return Path(raw).expanduser()
        return self.watch_dir() / "Fatura_Takip_Liste.xlsx"

    def _refresh_mail_label(self):
        if self.db.get_setting("mail_enabled") == "1":
            self.mail_state.configure(text=f"E-posta açık\n{self.db.get_setting('mail_user', '')}")
        else:
            self.mail_state.configure(text="E-posta kapalı\nAyarlardan açılır")

    def show_section(self, name):
        self.section = name
        self.heading.configure(text=PAGE_TITLES.get(name, name))
        hints = {
            "İş takibi": long_today() + "  ·  açık işler, faturalar ve bu ayın neti",
            "Faturalar": "Filtreleme ve Sıralama menüsünden listeyi daralt veya sırala.",
            "Raporlar": "Kaydedilen harcama raporları.",
            "Görevler": "Durum başlığına basıp bir durum seç. Anahtar projeye bağlıdır.",
            "Hatırlatmalar": "Vakti gelince sağ altta uyarı çıkar ve ses çalar.",
            "Firmalar": "Yazdığın isim, açık olan Outlook adres defterinde aranır.",
            "Gelir-Gider": "Kayıt aç, fatura ekle. Sonradan kalemle belge ekleyip çıkarabilirsin.",
            "Projeler": "Proje seç, teklif ekle. PDF, e-posta, görüntü, Word ve Excel karşılaştırılır.",
        }
        self.subhead.configure(text=hints.get(name, ""))
        for key, button in self.nav_buttons.items():
            button.set_active(key == name)
        self.board.pack_forget()
        self.invoice_page.pack_forget()
        self.report_page.pack_forget()
        self.list_wrap.pack_forget()
        self.search_box.pack_forget()
        self.contact_find.pack_forget()
        self.offer_button.pack_forget()
        self.offer_panel.pack_forget()
        if name == "İş takibi":
            self.board.pack(fill="both", expand=True)
            self.refresh_board()
        elif name == "Faturalar":
            self.search_box.pack(side="right")
            self.invoice_page.pack(fill="both", expand=True)
            self.refresh_invoices()
        elif name == "Raporlar":
            self.search_box.pack(side="right")
            self.report_page.pack(fill="both", expand=True)
            self.refresh_reports()
        else:
            self.search_box.pack(side="right")
            self.list_wrap.pack(fill="both", expand=True)
            if name == "Firmalar":
                self.contact_find.pack(fill="x", padx=22, pady=(0, 8), before=self.table_frame)
            if name == "Projeler":
                self.offer_button.pack(side="left", padx=(6, 0))
                self.offer_panel.pack(fill="x", padx=22, pady=(8, 0), before=self.detail)
            self.refresh_list()

    def _on_search_changed(self, *_):
        job = getattr(self, "_search_job", None)
        if job is not None:
            self.after_cancel(job)
        delay = 450 if self.section == "Firmalar" else 40
        self._search_job = self.after(delay, self._run_search)

    def _run_search(self):
        self._search_job = None
        query = self.search_var.get().strip()
        if self.section == "Firmalar" and len(query) >= 2:
            self._address_token = getattr(self, "_address_token", 0) + 1
            token = self._address_token
            self.status.configure(text="Outlook adres defteri aranıyor…")
            threading.Thread(target=self._fetch_address_book, args=(token, query), daemon=True).start()
            return
        self._address_hits = []
        self._address_error = ""
        self.refresh_current()

    def _fetch_address_book(self, token, query):
        hits, error = lookup_address_book(query)
        self.after(0, lambda: self._apply_address_book(token, hits, error))

    def _apply_address_book(self, token, hits, error):
        if not self._alive or token != getattr(self, "_address_token", None):
            return
        self._address_hits = hits
        self._address_error = error
        if self.section == "Firmalar":
            self.refresh_list()

    def _address_hit_for(self, iid):
        if not str(iid).startswith("ext-"):
            return None
        try:
            index = int(str(iid).split("-", 1)[1])
        except ValueError:
            return None
        hits = getattr(self, "_address_hits", [])
        if index < 0 or index >= len(hits):
            return None
        return hits[index]

    def _save_address_hit(self, iid):
        hit = self._address_hit_for(iid)
        if hit is None:
            return

        class HitRow:
            def __init__(self, data):
                self._data = data

            def keys(self):
                return self._data.keys()

            def __getitem__(self, key):
                return self._data.get(key, "")

        seeded = HitRow({
            "record_kind": "Firma" if hit["company"] and not hit["name"] else "Kişi",
            "scope": "Şirket dışı",
            "name": hit["name"] or hit["company"],
            "contact_person": hit["company"],
            "phone": hit["phone"],
            "email": hit["email"],
            "notes": f"{hit['source']} adres defterinden alındı.",
        })
        dialog = ContactDialog(self, seeded)
        self.wait_window(dialog)
        if dialog.result:
            self.db.save("contacts", dialog.result)
            self.refresh_current()

    def refresh_current(self):
        if not self._alive or not hasattr(self, "tree"):
            return
        if self.section == "İş takibi":
            self.refresh_board()
        elif self.section == "Faturalar":
            self.refresh_invoices()
        elif self.section == "Raporlar":
            self.refresh_reports()
        else:
            self.refresh_list()

    def schema(self):
        return SCHEMAS[self.section]

    def refresh_board(self):
        tasks = [row for row in self.db.all("tasks") if row["status"] not in ("İptal",)]
        invoices = self.db.all("invoices")
        transactions = self.db.all("transactions")
        today_date = date.today()
        month = today_date.strftime("%Y-%m")
        open_tasks = [row for row in tasks if row["status"] != "Tamamlandı"]
        late = [row for row in open_tasks if (day := parse_day(row["due_date"])) and day < today_date]
        unpaid = [row for row in invoices if row["status"] in ("Ödenmedi", "Kısmi ödendi")]
        income = sum(float(row["amount"] or 0) for row in transactions if row["kind"] == "Gelir" and str(row["trans_date"] or "").startswith(month))
        expense = sum(float(row["amount"] or 0) for row in transactions if row["kind"] == "Gider" and str(row["trans_date"] or "").startswith(month))
        unpaid_total = sum(float(row["amount"] or 0) for row in unpaid)
        self.metric_values["open"].configure(text=str(len(open_tasks)), fg=INK)
        self.metric_notes["open"].configure(text=f"{len(tasks)} kayıt", fg=MUTED)
        self.metric_values["late"].configure(text=str(len(late)), fg=INK)
        self.metric_notes["late"].configure(text="vadesi geçmiş", fg=DANGER if late else MUTED)
        self.metric_values["unpaid"].configure(text=(format_amount(unpaid_total) or "0,00") + " ₺", fg=INK)
        self.metric_notes["unpaid"].configure(text=f"{len(unpaid)} fatura", fg=WARN if unpaid else MUTED)
        net = income - expense
        self.metric_values["net"].configure(text=(format_amount(net) or "0,00") + " ₺", fg=INK)
        self.metric_notes["net"].configure(text="gelir − gider", fg=OK if net >= 0 else DANGER)
        grouped = {"Bekliyor": [], "Devam ediyor": [], "Tamamlandı": []}
        for row in tasks:
            grouped.get(row["status"], grouped["Bekliyor"]).append(row)
        for name, holder in self.columns.items():
            for child in holder.winfo_children():
                child.destroy()
            def board_key(row, column=name):
                start = task_waiting_since(row) if column == "Bekliyor" else None
                breached = start is not None and datetime.now() >= start + timedelta(hours=24)
                return (0 if breached else 1, PRIORITY_RANK.get(row["priority"], 9), row["due_date"] or "9999")

            rows = sorted(grouped[name], key=board_key)
            if name in self.column_counts:
                self.column_counts[name].configure(text=str(len(rows)))
            if not rows:
                tk.Label(holder, text="Kayıt yok", bg=holder["bg"], fg=MUTED, font=(FONT, 12)).pack(anchor="w", padx=10, pady=14)
            for row in rows[:12]:
                self._task_card(holder, row, today_date)
        self._fill_recent(invoices)
        self.status.configure(text=f"Klasör: {self.watch_dir()}    ·    Excel: {self.excel_path().name}")

    def _task_card(self, parent, row, today_date):
        card = tk.Frame(parent, bg=CARD, highlightbackground=LINE, highlightthickness=1, cursor="hand2")
        card.pack(fill="x", pady=(0, 8))
        late = (day := parse_day(row["due_date"])) and day < today_date and row["status"] != "Tamamlandı"
        priority = row["priority"] or "Normal"
        status_color = {"Bekliyor": WARN, "Devam ediyor": ACCENT, "Tamamlandı": OK}.get(row["status"], MUTED)
        mark = DANGER if late or priority == "Acil" else WARN if priority == "Yüksek" else status_color
        tk.Frame(card, bg=mark, width=5).pack(side="left", fill="y")
        body = tk.Frame(card, bg=CARD, padx=10, pady=8)
        body.pack(side="left", fill="both", expand=True)
        title = tk.Label(body, text=row["title"], bg=CARD, fg=INK, font=(FONT, 13, "bold"), anchor="w", wraplength=220, justify="left")
        title.pack(fill="x")
        due = f"{day.day} {MONTHS[day.month - 1]}" if day else "Tarih yok"
        bits = []
        if priority in ("Acil", "Yüksek"):
            bits.append(priority)
        bits.append(due)
        bits.append(row["assignee"] or "Atanmadı")
        meta = tk.Label(body, text="  ·  ".join(bits), bg=CARD, fg=DANGER if late else MUTED, font=(FONT, 11), anchor="w", wraplength=220, justify="left")
        meta.pack(fill="x", pady=(3, 0))
        widgets = [card, body, title, meta]
        sla = self._sla_caption(row)
        if sla:
            sla_label = tk.Label(body, text=sla, bg=CARD, fg=DANGER, font=(FONT, 11), anchor="w", wraplength=220, justify="left")
            sla_label.pack(fill="x", pady=(3, 0))
            widgets.append(sla_label)
        if row["status"] in ("Bekliyor", "Devam ediyor") and row["id"] in getattr(self, "_board_tag_ids", set()):
            label = (row["tag"] or "").strip() or "Etiket yok"
            chip_bg = status_color if row["tag"] else "#e8edf2"
            chip_fg = "white" if row["tag"] else MUTED
            chip = tk.Label(body, text=label, bg=chip_bg, fg=chip_fg, font=(FONT, 11, "bold"), padx=8, pady=2)
            chip.pack(anchor="w", pady=(6, 0))
            widgets.append(chip)
        handler = lambda _e, row_id=row["id"], status=row["status"]: self._board_task_click(row_id, status)
        for widget in widgets:
            widget.bind("<Button-1>", handler)

    def _board_task_click(self, row_id, status):
        if status in ("Bekliyor", "Devam ediyor"):
            shown = getattr(self, "_board_tag_ids", set())
            shown.add(row_id)
            self._board_tag_ids = shown
            self.refresh_board()
            return
        self._edit_task(row_id)

    def _sla_caption(self, row):
        if row["status"] in ("Tamamlandı", "İptal"):
            return ""
        start = task_waiting_since(row)
        if start is None or datetime.now() < start + timedelta(hours=24):
            return ""
        moment = task_sla_moment(row)
        if moment is not None and moment > datetime.now():
            clock = moment.strftime("%H:%M") if moment.date() == date.today() else moment.strftime("%d.%m %H:%M")
            return f"SLA doldu · {clock} tarihine ertelendi"
        return "SLA doldu"

    def _save_task(self, data, row_id=None, previous=None):
        data = dict(data)
        self.db.ensure_workspace()
        if "project_label" in data:
            project = self._project_from_label(data.pop("project_label")) or self._current_board_project()
            data["project_id"] = project["id"]
        elif not data.get("project_id"):
            if previous is not None and "project_id" in previous.keys() and previous["project_id"]:
                data["project_id"] = previous["project_id"]
            else:
                data["project_id"] = self._current_board_project()["id"]
        if not str(data.get("issue_type") or "").strip():
            if previous is not None and "issue_type" in previous.keys() and previous["issue_type"]:
                data["issue_type"] = previous["issue_type"]
            else:
                data["issue_type"] = "Görev"
        previous_project = previous["project_id"] if previous is not None and "project_id" in previous.keys() else None
        previous_no = previous["issue_no"] if previous is not None and "issue_no" in previous.keys() else None
        if not previous_no or previous_project != data["project_id"]:
            data["issue_no"] = self.db.next_issue_no(data["project_id"])
        stamp_task_wait(data, previous)
        return self.db.save("tasks", data, row_id)

    def _edit_task(self, row_id):
        row = self.db.get("tasks", row_id)
        if row is None:
            return
        self._open_task_dialog(row)

    def _sorted_projects(self):
        self.db.ensure_workspace()
        return sorted(self.db.all("projects"), key=lambda project: fold(project["name"] or ""))

    def _project_label(self, project):
        key = project["proj_key"] or "PRJ"
        return f"{key}  ·  {project['name']}"

    def _project_from_label(self, label):
        for project in self.db.all("projects"):
            if self._project_label(project) == label:
                return project
        return None

    def _project_by_id(self, project_id):
        if not project_id:
            return None
        try:
            return self.db.get("projects", int(project_id))
        except (TypeError, ValueError):
            return None

    def _project_name(self, row):
        project = self._project_by_id(row["project_id"] if "project_id" in row.keys() else None)
        if project is None or not project["name"]:
            return "(boş)"
        return project["name"]

    def _current_board_project(self):
        projects = self._sorted_projects()
        wanted = getattr(self, "board_project_id", None)
        if wanted is None:
            saved = self.db.get_setting("board_project", "")
            if str(saved).isdigit():
                wanted = int(saved)
        match = self._project_by_id(wanted) if wanted else None
        if match is None:
            match = next((project for project in projects if (project["proj_key"] or "").upper() == "GEN"), projects[0])
            self.board_project_id = match["id"]
        return match

    def issue_key(self, row):
        number = row["issue_no"] if "issue_no" in row.keys() and row["issue_no"] else row["id"]
        project = self._project_by_id(row["project_id"] if "project_id" in row.keys() else None)
        prefix = project["proj_key"] if project is not None and project["proj_key"] else "PRJ"
        return f"{prefix}-{int(number)}"

    def _save_project(self, data, row_id=None):
        data = dict(data)
        current = self.db.get("projects", row_id) if row_id else None
        existing = ""
        if current is not None and "proj_key" in current.keys() and current["proj_key"]:
            existing = "".join(ch for ch in current["proj_key"].upper() if ch.isalnum())[:6]
        typed = "".join(ch for ch in str(data.get("proj_key") or "").upper() if ch.isalnum())[:6]
        taken = set()
        for project in self.db.all("projects"):
            if row_id and project["id"] == int(row_id):
                continue
            if project["proj_key"]:
                taken.add("".join(ch for ch in project["proj_key"].upper() if ch.isalnum()))
        if typed and typed not in taken:
            data["proj_key"] = typed
        elif existing:
            data["proj_key"] = existing
        else:
            data["proj_key"] = project_key(data.get("name") or "PRJ", taken)
        self.db._workspace_ready = True
        return self.db.save("projects", data, row_id)

    def _task_initial(self, row, status, project_id):
        data = {}
        for _label, key, _kind, *_rest in SCHEMAS["Görevler"]["fields"]:
            if row is not None and key in row.keys() and row[key] is not None:
                data[key] = row[key]
            elif key == "issue_type":
                data[key] = "Görev"
            elif key == "priority":
                data[key] = "Normal"
            elif key == "status":
                data[key] = "Bekliyor"
            elif key == "repeat":
                data[key] = "Yok"
            else:
                data[key] = ""
        if status:
            data["status"] = status
        project = self._project_by_id(project_id) if project_id else None
        if project is None and row is not None and "project_id" in row.keys():
            project = self._project_by_id(row["project_id"])
        if project is None:
            project = self._current_board_project()
        data["project_label"] = self._project_label(project)
        return data

    def _open_task_dialog(self, row=None, status=None, project_id=None):
        self.db.ensure_workspace()
        labels = [self._project_label(project) for project in self._sorted_projects()]
        config = {
            "table": "tasks",
            "title": "Görev",
            "fields": [("Proje", "project_label", "choice", labels), *SCHEMAS["Görevler"]["fields"]],
        }
        dialog = RecordDialog(self, config, self._task_initial(row, status, project_id), self.watch_dir())
        self.wait_window(dialog)
        if not dialog.result:
            return None
        saved = self._save_task(dialog.result, None if row is None else row["id"], row)
        self.refresh_current()
        return saved


    def _wheel_recent(self, event):
        self.recent_canvas.yview_scroll(-1 if event.delta > 0 else 1, "units")

    def _bind_recent_wheel(self, widget):
        widget.bind("<MouseWheel>", self._wheel_recent)
        for child in widget.winfo_children():
            self._bind_recent_wheel(child)

    def _fill_recent(self, invoices):
        for child in self.recent.winfo_children():
            child.destroy()
        tk.Label(self.recent, text="Son faturalar", bg=CARD, fg=INK, font=(FONT, 12, "bold")).pack(anchor="w", pady=(8, 6))
        rows = list(invoices)
        chosen = None
        if self.selected_invoice:
            chosen = self.db.get("invoices", self.selected_invoice)
        if chosen is None and rows:
            chosen = rows[0]
            self.selected_invoice = chosen["id"]
        for row in rows:
            selected = chosen is not None and row["id"] == chosen["id"]
            bg = SOFT if selected else CARD
            item = tk.Frame(self.recent, bg=bg, cursor="hand2")
            item.pack(fill="x", pady=1)
            amount = format_amount(row["amount"])
            title = tk.Label(item, text=row["party"], bg=bg, fg=INK, font=(FONT, 12, "bold"), anchor="w", wraplength=300, justify="left", cursor="hand2")
            title.pack(fill="x", padx=8, pady=(6, 0))
            detail = "  ·  ".join(part for part in (f"{amount} ₺" if amount else "Tutar bekleniyor", row["invoice_no"] or "") if part)
            meta = tk.Label(item, text=detail, bg=bg, fg=MUTED, font=(FONT, 11), anchor="w", cursor="hand2")
            meta.pack(fill="x", padx=8, pady=(1, 6))
            handler = lambda _e, row_id=row["id"]: self.select_invoice(row_id, on_board=True)
            for widget in (item, title, meta):
                widget.bind("<Button-1>", handler)
        self._bind_recent_wheel(self.recent)
        self.recent.update_idletasks()
        self.recent_canvas.configure(scrollregion=self.recent_canvas.bbox("all"))
        self.board_preview.show(chosen)

    def current_invoices(self):
        query = fold(self.search_var.get().strip())
        mode = self.invoice_filter.get()
        rows = []
        for row in self.db.all("invoices"):
            if mode == "Ödenmedi" and row["status"] not in ("Ödenmedi", "Kısmi ödendi"):
                continue
            if mode == "Ödendi" and row["status"] != "Ödendi":
                continue
            if mode == "Klasör" and row["source"] != "klasor":
                continue
            if mode == "E-posta" and row["source"] != "eposta":
                continue
            source = {"klasor": "Klasör", "eposta": "E-posta", "manuel": "Elle"}.get(row["source"] or "manuel", "Elle")
            blob = " ".join(fold(part) for part in (
                row["party"], row["invoice_no"], row["status"], row["description"], row["file_name"],
                row["invoice_date"], row["due_date"], row["excel_status"], row["tax_no"], row["report_reason"], source, amount_terms(row["amount"]),
            ))
            if query and query not in blob:
                continue
            rows.append(row)
        return self._sort_invoices(rows)

    def _sort_invoices(self, rows):
        kind = self.invoice_sort.get()

        def day_of(row):
            return parse_day(row["invoice_date"]) or parse_day(row["due_date"])

        if kind == "Tutar (yüksek)":
            return sorted(rows, key=lambda row: float(row["amount"]) if row["amount"] not in (None, "") else -1, reverse=True)
        if kind == "Tutar (düşük)":
            return sorted(rows, key=lambda row: (row["amount"] in (None, ""), float(row["amount"] or 0)))
        if kind == "Tarih (eski)":
            return sorted(rows, key=lambda row: (day_of(row) is None, day_of(row) or date.max))
        if kind == "Firma":
            return sorted(rows, key=lambda row: str(row["party"] or "").casefold())
        return sorted(rows, key=lambda row: day_of(row) or date.min, reverse=True)

    def _wheel_invoices(self, event):
        self.invoice_canvas.yview_scroll(-1 if event.delta > 0 else 1, "units")

    def _bind_invoice_wheel(self, widget):
        widget.bind("<MouseWheel>", self._wheel_invoices)
        for child in widget.winfo_children():
            self._bind_invoice_wheel(child)

    def refresh_invoices(self):
        if not hasattr(self, "invoice_list"):
            return
        place = self.invoice_canvas.yview()[0] if self.invoice_list.winfo_children() else 0
        for child in self.invoice_list.winfo_children():
            child.destroy()
        rows = self.current_invoices()
        if not rows:
            empty = "Aramaya uyan fatura yok." if self.search_var.get().strip() else "Bu görünümde fatura yok."
            tk.Label(self.invoice_list, text=empty, bg=CANVAS, fg=MUTED, font=(FONT, 13)).pack(anchor="w", pady=12)
        chosen = None
        for row in rows:
            selected = row["id"] == self.selected_invoice
            if selected:
                chosen = row
            self._invoice_card(row, selected)
        if chosen is None and rows:
            chosen = rows[0]
            self.selected_invoice = chosen["id"]
            self.refresh_invoices()
            return
        self._bind_invoice_wheel(self.invoice_list)
        self.invoice_list.update_idletasks()
        self.invoice_canvas.configure(scrollregion=self.invoice_canvas.bbox("all"))
        self.invoice_canvas.yview_moveto(place)
        self.invoice_preview.show(chosen)
        self.status.configure(text=f"{len(rows)} fatura")

    def _invoice_card(self, row, selected):
        card = tk.Frame(self.invoice_list, bg=CARD, highlightbackground=ACCENT if selected else LINE, highlightthickness=2 if selected else 1, padx=12, pady=8, cursor="hand2")
        card.pack(fill="x", pady=(0, 8))
        top = tk.Frame(card, bg=CARD)
        top.pack(fill="x")
        open_row = getattr(self, "expanded_invoice", None) == row["id"]
        arrow = tk.Label(top, text="▲" if open_row else "▼", bg=CARD, fg=ACCENT if open_row else MUTED, font=(FONT, 16), cursor="hand2", padx=6)
        arrow.pack(side="right")
        amount = format_amount(row["amount"])
        amount_label = tk.Label(top, text=(amount + " ₺") if amount else "—", bg=CARD, fg=INK, font=(FONT, 14, "bold"))
        amount_label.pack(side="right")
        party = tk.Label(top, text=row["party"], bg=CARD, fg=INK, font=(FONT, 14, "bold"), wraplength=420, justify="left", anchor="w")
        party.pack(side="left", fill="x", expand=True)
        pill_bg, pill_fg = STATUS_PILL.get(row["status"], ("#eee", INK))
        source = {"klasor": "Klasör", "eposta": "E-posta"}.get(row["source"] or "", "Elle")
        bottom = tk.Frame(card, bg=CARD)
        bottom.pack(fill="x", pady=(4, 0))
        status = tk.Label(bottom, text=row["status"] or "", bg=pill_bg, fg=pill_fg, font=(FONT, 10, "bold"), padx=6, pady=1)
        status.pack(side="left")
        source_label = tk.Label(bottom, text=source, bg=SOFT, fg=OK, font=(FONT, 10, "bold"), padx=6, pady=1)
        source_label.pack(side="left", padx=6)
        meta = tk.Label(bottom, text=f"{row['invoice_no'] or 'no yok'}   ·   {row['due_date'] or row['invoice_date'] or ''}", bg=CARD, fg=MUTED, font=(FONT, 11))
        meta.pack(side="left", padx=4)
        due_chip = None
        due = parse_day(row["due_date"])
        if due and row["status"] in ("Ödenmedi", "Kısmi ödendi"):
            delta = (due - date.today()).days
            if delta < 0:
                chip_text, chip_bg, chip_fg = f"{abs(delta)} gün gecikti", DANGER_SOFT, DANGER
            elif delta == 0:
                chip_text, chip_bg, chip_fg = "Bugün", WARN_SOFT, WARN
            else:
                chip_text, chip_bg, chip_fg = f"{delta} gün", WARN_SOFT, WARN
            due_chip = tk.Label(bottom, text=chip_text, bg=chip_bg, fg=chip_fg, font=(FONT, 10, "bold"), padx=6, pady=1)
            due_chip.pack(side="right")
        if open_row:
            actions = tk.Frame(card, bg=CARD)
            actions.pack(fill="x", pady=(8, 0))
            Click(actions, "Düzenle", lambda row_id=row["id"]: self._edit_invoice(row_id), ACCENT, "white", ACCENT_HOVER, font=(FONT, 12, "bold"), padx=14, pady=6).pack(side="left")
        handler = lambda _e, row_id=row["id"]: self.select_invoice(row_id, on_board=False)
        arrow.bind("<Button-1>", lambda _e, row_id=row["id"]: self.toggle_invoice_edit(row_id))
        for widget in (card, top, party, amount_label, bottom, status, source_label, meta):
            widget.bind("<Button-1>", handler)
        if due_chip is not None:
            due_chip.bind("<Button-1>", handler)

    def toggle_invoice_edit(self, row_id):
        self.selected_invoice = row_id
        self.expanded_invoice = None if getattr(self, "expanded_invoice", None) == row_id else row_id
        self.refresh_invoices()
        return "break"

    def open_invoice(self, row_id):
        if getattr(self, "_opening_invoice", False):
            return
        self._opening_invoice = True
        self.selected_invoice = row_id
        self.after(30, lambda rid=row_id: self._open_invoice_editor(rid))

    def _open_invoice_editor(self, row_id):
        try:
            if not self._alive:
                return
            self.refresh_invoices()
            self._edit_invoice(row_id)
        finally:
            self._opening_invoice = False

    def select_invoice(self, row_id, on_board):
        self.selected_invoice = row_id
        row = self.db.get("invoices", row_id)
        if on_board:
            self._fill_recent(self.db.all("invoices"))
        else:
            self.refresh_invoices()
        if row and self.section == "İş takibi":
            self.board_preview.show(row)

    def visible_rows(self):
        config = self.schema()
        query = fold(self.search_var.get().strip())
        rows = self.db.all(config["table"])
        keys = [key for key, _label in config["columns"]]

        def matches(row):
            parts = []
            for key in keys:
                if key in row.keys():
                    parts.append(str(row[key] or ""))
                else:
                    parts.append(self._cell_display(row, key))
            blob = fold(" ".join(parts))
            if query in blob:
                return True
            if self.section == "Görevler":
                raw = row["note_json"] if "note_json" in row.keys() else ""
                if query in fold(note_plain(raw)):
                    return True
            if self.section == "Gelir-Gider":
                raw = row["docs_json"] if "docs_json" in row.keys() else ""
                names = " ".join(item["name"] for item in ledger_documents(raw))
                if query in fold(names):
                    return True
            if self.section == "Firmalar":
                fields = ("name", "contact_person", "email", "phone", "extension", "address", "notes", "tax_no", "record_kind", "scope")
                blob = fold(" ".join(str(row[key] or "") for key in fields if key in row.keys()))
                phones = "".join(ch for ch in blob if ch.isdigit())
                return all(word in blob or (word.isdigit() and word in phones) for word in query.split())
            if self.section != "Projeler":
                return False
            for offer in self.db.offers_for(row["id"]):
                extra = fold(" ".join(str(offer[key] or "") for key in ("vendor", "title", "pros", "cons", "notes")))
                if query in extra:
                    return True
            return False

        if query:
            rows = [row for row in rows if matches(row)]
        filters = self._active_column_filters()
        for key, allowed in filters.items():
            rows = [row for row in rows if self._cell_display(row, key) in allowed]
        return rows

    def refresh_list(self):
        if self.section in ("İş takibi", "Faturalar", "Raporlar") or not hasattr(self, "tree"):
            return
        config = self.schema()
        columns = []
        for key, label in config["columns"]:
            columns.append((key, label))
            if self.section == "Görevler" and key == "status":
                columns.append(("_note", "Yorum"))
        if self.section == "Firmalar":
            columns.append(("_edit", ""))
        if self.section == "Gelir-Gider":
            columns.append(("_docs", "Belge"))
            columns.append(("_edit", ""))
        if self.section == "Projeler":
            columns.append(("offer_count", "Teklif"))
        keep = self.tree.selection()
        self._paint_text = {}
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = [key for key, _label in columns]
        text_font = tkfont.Font(self, font=(FONT, 13))
        head_font = tkfont.Font(self, font=(FONT, 12, "bold"))
        fitted = {}
        for key, label in columns:
            picked = self._active_column_filters().get(key)
            if key == "_edit":
                title = label
            elif picked and len(picked) == 1:
                title = f"{label}: {next(iter(picked))}  ▾"
            else:
                title = f"{label}  ▾"
            if key in {"_edit", "_note", "offer_count", "_docs"}:
                anchor = "center"
                fixed = {"_edit": 52, "_note": 88, "offer_count": 110, "_docs": 100}[key]
                width = max(fixed, head_font.measure(title) + 28)
                minimum = width
                stretch = False
            else:
                anchor = "e" if key in MONEY_KEYS or key in {"due_date", "invoice_date", "start_date", "end_date", "trans_date", "remind_date"} else "w"
                width = head_font.measure(title) + 36
                minimum = width
                stretch = False
                fitted[key] = width
            self.tree.heading(key, text=title, anchor=anchor)
            self.tree.column(key, width=width, minwidth=minimum, anchor=anchor, stretch=stretch)
        rows = self.visible_rows()
        for index, row in enumerate(rows):
            values = []
            for key, _label in columns:
                if key == "_note":
                    raw = row["note_json"] if "note_json" in row.keys() else ""
                    values.append("●" if note_has_items(raw) else "+")
                elif key == "_docs":
                    raw = row["docs_json"] if "docs_json" in row.keys() else ""
                    count = len(ledger_documents(raw))
                    values.append(f"{count} belge" if count else "—")
                elif key == "_edit":
                    values.append("" if self.section == "Gelir-Gider" else "✎")
                elif key == "offer_count":
                    count = self.db.offer_count(row["id"])
                    values.append(f"{count} teklif" if count else "—")
                elif key in MONEY_KEYS:
                    values.append(format_amount(row[key]))
                elif key not in row.keys():
                    values.append(self._cell_display(row, key))
                else:
                    values.append("" if row[key] is None else row[key])
                if key in fitted:
                    fitted[key] = max(fitted[key], text_font.measure(str(values[-1])) + 36)
            status = row["status"] if "status" in row.keys() else ""
            tag = {"Tamamlandı": "done", "Ödendi": "done", "Devam ediyor": "run", "Planlandı": "wait", "Beklemede": "wait", "Bekliyor": "wait", "İptal": "odd"}.get(status, "odd" if index % 2 else "even")
            if self.section == "Görevler":
                tag = "plain"
                keys = [key for key, _label in columns]
                stored = {}
                for key in ("title", "priority"):
                    if key in keys:
                        stored[key] = values[keys.index(key)]
                        values[keys.index(key)] = ""
                self._paint_text[str(row["id"])] = stored
                for key, text in stored.items():
                    if key in fitted:
                        fitted[key] = max(fitted[key], text_font.measure(str(text)) + 36)
            self.tree.insert("", "end", iid=str(row["id"]), values=values, tags=(tag,))
        book_hits = []
        if self.section == "Firmalar" and len(self.search_var.get().strip()) >= 2:
            known_mails = {fold(row["email"]) for row in rows if "email" in row.keys() and row["email"]}
            known_names = {fold(row["name"]) for row in rows if "name" in row.keys() and row["name"]}
            for hit in getattr(self, "_address_hits", []):
                mail = fold(hit.get("email"))
                if mail and mail in known_mails:
                    continue
                if not mail and fold(hit.get("name")) in known_names:
                    continue
                book_hits.append(hit)
        for index, hit in enumerate(book_hits):
            values = []
            for key, _label in columns:
                text = {
                    "record_kind": hit["source"],
                    "scope": "Adres defteri",
                    "name": hit["name"],
                    "contact_person": hit["company"],
                    "phone": hit["phone"],
                    "email": hit["email"],
                    "_edit": "✎",
                }.get(key, "")
                values.append(text)
                if key in fitted:
                    fitted[key] = max(fitted[key], text_font.measure(str(text)) + 36)
            self.tree.insert("", "end", iid=f"ext-{index}", values=values, tags=("book",))
        for key, width in fitted.items():
            width = max(72, min(width, 520))
            self.tree.column(key, width=width, minwidth=width)
        self.detail.configure(text="")
        note = f"{len(rows) + len(book_hits)} kayıt"
        active = self._active_column_filters()
        if active:
            names = dict(self.schema()["columns"])
            names.update({"_note": "Yorum", "_docs": "Belge", "offer_count": "Teklif"})
            note += "  ·  süzülmüş: " + ", ".join(names.get(key, key) for key in active)
        if book_hits:
            note += f"  ·  adres defterinden {len(book_hits)}"
        if self.section == "Firmalar" and getattr(self, "_address_error", ""):
            note += "  ·  " + self._address_error
        self.status.configure(text=note)
        if rows or book_hits:
            self.empty_wrap.place_forget()
        else:
            if self._active_column_filters():
                self.empty_title.configure(text="Bu filtreye uyan kayıt yok")
                self.empty_note.configure(text="Sütun başlığındaki oka basıp Tümü’nü seç.")
            elif self.search_var.get().strip():
                self.empty_title.configure(text="Sonuç yok")
                self.empty_note.configure(text="Aramayı silince liste geri gelir.")
            else:
                titles = {
                    "Görevler": ("Henüz görev yok", "İlk işi eklediğinde burada görünür."),
                    "Hatırlatmalar": ("Hatırlatma yok", "Tarihi takvimden seç, saati HH:MM yaz. Vakti gelince sağ altta uyarı çıkar."),
                    "Firmalar": ("Rehber boş", "Yeni kayıtta önce kişi veya firma, sonra şirket içi ya da dışı seçilir."),
                    "Gelir-Gider": ("Kayıt yok", "Yeni kayıtta fatura ekleyebilirsin. Sonradan satıra basınca çıkan kalemle belge eklenir veya çıkarılır."),
                    "Projeler": ("Henüz proje yok", "Projeyi ekle, sonra Teklif ekle ile PDF veya kayıtlı e-postayı karşılaştır."),
                }
                title, note = titles.get(self.section, ("Kayıt yok", ""))
                self.empty_title.configure(text=title)
                self.empty_note.configure(text=note)
            self.empty_wrap.place(relx=0.5, rely=0.46, anchor="center")
            self.empty_wrap.lift()
        if self.section == "Projeler" and keep and self.tree.exists(keep[0]):
            self.tree.selection_set(keep[0])
        elif self.section == "Projeler":
            self.refresh_offers()
        self._schedule_cell_paint()
        self._mark_ledger_pencil()

    def _on_tree_yview(self, first, last):
        self.tree_scroll.set(first, last)
        self._schedule_cell_paint()

    def _on_tree_select(self, _event):
        self.show_detail()
        self._schedule_cell_paint()
        self._mark_ledger_pencil()

    def _mark_ledger_pencil(self):
        if self.section != "Gelir-Gider" or "_edit" not in self.tree["columns"]:
            return
        selected = set(self.tree.selection())
        for item in self.tree.get_children():
            self.tree.set(item, "_edit", "✎" if item in selected else "")

    def _schedule_cell_paint(self):
        token = getattr(self, "_paint_after", None)
        if token:
            self.after_cancel(token)
        self._paint_after = self.after(16, self._paint_task_cells)

    def _cell_label(self, item, key):
        label = self._cell_labels.get((item, key))
        if label is not None and label.winfo_exists():
            return label
        label = tk.Label(self.tree, font=(FONT, 13), anchor="w", padx=4, borderwidth=0, highlightthickness=0)
        label.bind("<Button-1>", lambda _e, iid=item: self._pick_painted_row(iid))
        label.bind("<Double-Button-1>", lambda _e, iid=item: self._pick_painted_row(iid, edit=True))
        self._cell_labels[(item, key)] = label
        return label

    def _pick_painted_row(self, item, edit=False):
        if self.tree.exists(item):
            self.tree.selection_set(item)
            self.tree.focus(item)
        if edit:
            self.after(10, self.edit_record)
        return "break"

    def _paint_task_cells(self):
        self._paint_after = None
        pool = getattr(self, "_cell_labels", {})
        self._cell_labels = pool
        if self.section != "Görevler" or not hasattr(self, "tree"):
            for label in pool.values():
                label.destroy()
            self._cell_labels = {}
            return
        children = self.tree.get_children()
        if children and all(not self.tree.bbox(item, "title") for item in children[:6]):
            tries = getattr(self, "_paint_tries", 0) + 1
            self._paint_tries = tries
            if tries < 8:
                self._schedule_cell_paint()
            return
        self._paint_tries = 0
        live = set()
        selected = set(self.tree.selection())
        for index, item in enumerate(children):
            stored = self._paint_text.get(item, {})
            for key in ("title", "priority"):
                box = self.tree.bbox(item, key)
                label = self._cell_label(item, key)
                live.add((item, key))
                if not box:
                    label.place_forget()
                    continue
                text = stored.get(key, "")
                if key == "title":
                    bg, fg = TASK_STRIPE[index % 2]
                    if item in selected:
                        bg = "#D9E4FF" if index % 2 else SOFT
                else:
                    bg, fg = PRIORITY_CELL.get(text, PRIORITY_CELL["Normal"])
                label.configure(text=text, bg=bg, fg=fg)
                label.place(x=box[0], y=box[1], width=box[2], height=box[3])
                label.lift()
        for token, label in list(pool.items()):
            if token not in live:
                label.destroy()
                del pool[token]

    def selected_row(self):
        if self.section == "İş takibi":
            return None
        if self.section == "Faturalar":
            if not self.selected_invoice:
                messagebox.showinfo(APP_NAME, "Önce bir fatura seç.")
                return None
            return self.db.get("invoices", self.selected_invoice)
        selected = self.tree.selection()
        if not selected:
            messagebox.showinfo(APP_NAME, "Önce listeden bir kayıt seç.")
            return None
        if self._address_hit_for(selected[0]):
            return None
        if not str(selected[0]).isdigit():
            return None
        return self.db.get(self.schema()["table"], int(selected[0]))

    def _column_key_at(self, x):
        token = self.tree.identify_column(x)
        if not token.startswith("#"):
            return ""
        try:
            index = int(token[1:]) - 1
        except ValueError:
            return ""
        columns = list(self.tree["columns"])
        if index < 0 or index >= len(columns):
            return ""
        return columns[index]

    def _on_tree_double(self, event):
        if self._column_key_at(event.x) in {"_note", "_edit", "_docs"}:
            return "break"
        self.edit_record()

    def _on_tree_motion(self, event):
        key = self._column_key_at(event.x)
        row_id = self.tree.identify_row(event.y)
        pencil = key == "_edit" and self.section == "Gelir-Gider" and row_id and self.tree.set(row_id, "_edit") == "✎"
        docs = key == "_docs" and self.section == "Gelir-Gider" and row_id and self.tree.set(row_id, "_docs") not in {"", "—"}
        over = bool(row_id) and (key in {"_note", "_edit"} and self.section != "Gelir-Gider" or pencil or key == "_note" or docs)
        self.tree.configure(cursor="hand2" if over else "")

    def _on_tree_click(self, event):
        region = ""
        try:
            region = self.tree.identify_region(event.x, event.y)
        except tk.TclError:
            region = ""
        heading = region == "heading" or (region != "separator" and not self.tree.identify_row(event.y) and event.y <= 40)
        if heading:
            key = self._column_key_at(event.x)
            if key and key != "_edit":
                labels = {column: title for column, title in self.schema()["columns"]}
                labels.update({"_note": "Yorum", "_docs": "Belge", "offer_count": "Teklif"})
                self.open_column_filter(key, labels.get(key, key))
            return "break"
        key = self._column_key_at(event.x)
        row_id = self.tree.identify_row(event.y)
        if key == "_docs" and self.section == "Gelir-Gider" and row_id and str(row_id).isdigit():
            self.tree.selection_set(row_id)
            if self.tree.set(row_id, "_docs") not in {"", "—"}:
                self.after(10, lambda rid=int(row_id): self._open_ledger_docs(rid))
            return "break"
        if key not in {"_note", "_edit"} or not row_id:
            return
        if key == "_edit" and self.section == "Gelir-Gider":
            shown = self.tree.set(row_id, "_edit") == "✎"
            self.tree.selection_set(row_id)
            if shown:
                self.after(10, self.edit_record)
            return "break"
        self.tree.selection_set(row_id)
        if key == "_edit" and self.section == "Firmalar":
            if self._address_hit_for(row_id):
                self.after(10, lambda iid=row_id: self._save_address_hit(iid))
            else:
                self.after(10, lambda rid=int(row_id): self._edit_contact(rid))
            return "break"
        if key == "_note" and self.section == "Görevler":
            self.after(10, lambda rid=int(row_id): self.open_task_note(rid))
            return "break"

    def _open_ledger_docs(self, row_id):
        row = self.db.get("transactions", row_id)
        if row is None:
            return
        raw = row["docs_json"] if "docs_json" in row.keys() else ""
        docs = []
        for item in ledger_documents(raw):
            try:
                resolved = assert_openable(item["path"], [LEDGER_DIR])
            except ValueError:
                continue
            docs.append((item["name"], resolved))
        if not docs:
            messagebox.showinfo(APP_NAME, "Bu kayıtta açılacak belge yok.")
            return
        if len(docs) == 1:
            open_path(docs[0][1])
            return
        menu = tk.Menu(self, tearoff=0)
        for name, path in docs:
            menu.add_command(label=name, command=lambda path=path: open_path(path))
        try:
            menu.tk_popup(self.winfo_pointerx(), self.winfo_pointery())
        finally:
            menu.grab_release()

    def open_task_note(self, row_id):
        if getattr(self, "_note_open", False):
            return
        row = self.db.get("tasks", row_id)
        if row is None:
            return
        raw = row["note_json"] if "note_json" in row.keys() else ""
        self._note_open = True
        try:
            dialog = TaskNoteDialog(self, row["title"], raw, row["id"])
            self.wait_window(dialog)
        finally:
            self._note_open = False
        if dialog.saved is None:
            return
        self.db.save("tasks", {"note_json": dialog.saved}, row["id"])
        self.refresh_current()

    def _edit_contact(self, row_id):
        row = self.db.get("contacts", row_id)
        if row is None:
            return
        dialog = ContactDialog(self, row)
        self.wait_window(dialog)
        if dialog.result:
            self.db.save("contacts", dialog.result, row["id"])
            self.refresh_current()

    def show_detail(self):
        selected = self.tree.selection()
        if not selected or self.section in ("İş takibi", "Faturalar"):
            return
        hit = self._address_hit_for(selected[0])
        if hit:
            parts = [hit["source"] + " adres defteri", hit["name"], hit["company"], hit["email"], hit["phone"]]
            self.detail.configure(text="  ·  ".join(part for part in parts if part))
            return
        if not str(selected[0]).isdigit():
            return
        row = self.db.get(self.schema()["table"], int(selected[0]))
        if row is None:
            return
        parts = [str(row[key]) for key in DETAIL_KEYS if key in row.keys() and row[key]]
        self.detail.configure(text="  ·  ".join(parts))
        if self.section == "Projeler":
            self.refresh_offers()

    def add_record(self):
        if self.section == "Raporlar":
            self.prepare_report()
            return
        if self.section == "İş takibi":
            self._open_task_dialog()
            return
        if self.section == "Firmalar":
            dialog = ContactDialog(self)
            self.wait_window(dialog)
            if dialog.result:
                self.db.save("contacts", dialog.result)
                self.refresh_current()
            return
        if self.section == "Görevler":
            self._open_task_dialog()
            return
        dialog = RecordDialog(self, self.schema(), watch_root=self.watch_dir())
        self.wait_window(dialog)
        if dialog.result:
            data = dialog.result
            if self.section == "Faturalar":
                data["source"] = "manuel"
            if self.section == "Projeler":
                self._save_project(data)
            else:
                self.db.save(self.schema()["table"], data)
            self.refresh_current()
            if self.section == "Projeler":
                newest = self.db.all("projects")
                if newest and self.tree.exists(str(newest[0]["id"])):
                    self.tree.selection_set(str(newest[0]["id"]))
                    self.tree.see(str(newest[0]["id"]))
            if self.section == "Hatırlatmalar":
                self.after(300, self._scan_due)

    def edit_record(self):
        if self.section == "Raporlar":
            return
        if self.section == "Firmalar" and self.tree.selection() and self._address_hit_for(self.tree.selection()[0]):
            self._save_address_hit(self.tree.selection()[0])
            return
        if self.section == "Faturalar" and self.selected_invoice:
            self._edit_invoice(self.selected_invoice)
            return
        row = self.selected_row()
        if row is None:
            return
        if self.section == "Firmalar":
            dialog = ContactDialog(self, row)
            self.wait_window(dialog)
            if dialog.result:
                self.db.save("contacts", dialog.result, row["id"])
                self.refresh_current()
            return
        if self.section == "Görevler":
            self._open_task_dialog(row)
            return
        dialog = RecordDialog(self, self.schema(), row, self.watch_dir())
        self.wait_window(dialog)
        if dialog.result:
            if self.section == "Projeler":
                self._save_project(dialog.result, row["id"])
            else:
                self.db.save(self.schema()["table"], dialog.result, row["id"])
            self.refresh_current()
            if self.section == "Projeler" and self.tree.exists(str(row["id"])):
                self.tree.selection_set(str(row["id"]))
            if self.section == "Hatırlatmalar":
                self.after(300, self._scan_due)

    def _warm_report_tool(self):
        try:
            report_tool()
            offer_tool()
            page_tool()
        except (OSError, ValueError, subprocess.SubprocessError):
            return

    def prepare_report(self, report=None):
        if not self.db.all("invoices"):
            messagebox.showinfo(APP_NAME, "Rapor için önce bir fatura olmalı.")
            return
        dialog = ReportDialog(self, report)
        self.wait_window(dialog)
        if not dialog.result:
            return
        saved = dialog.result
        row = dialog.row
        self.db.save("invoices", {"report_reason": saved["report_reason"], "report_files": saved["report_files"]}, row["id"])
        payload = {
            "invoice_id": row["id"],
            "title": row["party"] or "Fatura",
            "invoice_no": row["invoice_no"] or "",
            "amount": row["amount"],
            "reason": saved["report_reason"],
            "saved_path": saved["saved_path"],
        }
        if report is None:
            self.db.save("reports", payload)
        else:
            self.db.save("reports", payload, report["id"])
        self.status.configure(text=f"PDF kaydedildi: {saved['saved_path']}")
        if self.section != "Raporlar":
            self.show_section("Raporlar")
        else:
            self.refresh_reports()

    def refresh_reports(self):
        if not hasattr(self, "report_list"):
            return
        for child in self.report_list.winfo_children():
            child.destroy()
        query = fold(self.search_var.get().strip())
        rows = []
        for row in self.db.all("reports"):
            blob = fold(" ".join(str(row[key] or "") for key in ("title", "invoice_no", "reason", "saved_path")))
            if query and query not in blob:
                continue
            rows.append(row)
        if not rows:
            note = "Aramaya uyan rapor yok." if query else "Henüz rapor yok. Rapor Hazırla, faturayı, gerekçeyi ve ekleri tek PDF yapar."
            tk.Label(self.report_list, text=note, bg=CANVAS, fg=MUTED, font=(FONT, 13), wraplength=640, justify="left").pack(anchor="w", pady=12)
        opened = getattr(self, "expanded_report", None)
        for row in rows:
            selected = opened == row["id"]
            card = tk.Frame(self.report_list, bg=CARD, highlightbackground=ACCENT if selected else LINE, highlightthickness=1, padx=12, pady=10, cursor="hand2")
            card.pack(fill="x", pady=(0, 8))
            amount = format_amount(row["amount"])
            head = tk.Frame(card, bg=CARD)
            head.pack(fill="x")
            title = tk.Label(head, text=row["title"] or "Rapor", bg=CARD, fg=INK, font=(FONT, 14, "bold"), anchor="w", cursor="hand2")
            title.pack(side="left", fill="x", expand=True)
            if selected:
                Click(head, "🗑", lambda report_id=row["id"]: self._delete_report(report_id), CARD, DANGER, "#fdecea", font=(FONT, 14), padx=8, pady=2, border=LINE).pack(side="right")
                Click(head, "✎", lambda report_id=row["id"]: self._edit_report(report_id), CARD, INK, HOVER, font=(FONT, 14), padx=8, pady=2, border=LINE).pack(side="right", padx=(0, 6))
            meta = tk.Label(card, text=f"{amount + ' ₺' if amount else 'tutar yok'}   ·   {row['invoice_no'] or 'no yok'}", bg=CARD, fg=MUTED, font=(FONT, 12), anchor="w", cursor="hand2")
            meta.pack(fill="x", pady=(2, 4))
            reason = tk.Label(card, text=one_line(row["reason"], 180), bg=CARD, fg=INK, font=(FONT, 12), anchor="w", justify="left", wraplength=760, cursor="hand2")
            reason.pack(fill="x")
            path_label = tk.Label(card, text=row["saved_path"] or "", bg=CARD, fg=MUTED, font=(FONT, 11), anchor="w", cursor="hand2")
            path_label.pack(fill="x", pady=(4, 6))
            Click(card, "PDF'i aç", lambda path=row["saved_path"]: self._open_report(path), ACCENT, "white", ACCENT_HOVER, font=(FONT, 12, "bold"), padx=12, pady=6).pack(anchor="w")
            toggle = lambda _e, report_id=row["id"]: self._toggle_report(report_id)
            for widget in (card, head, title, meta, reason, path_label):
                widget.bind("<Button-1>", toggle)
        self.status.configure(text=f"{len(rows)} rapor")

    def _toggle_report(self, report_id):
        self.expanded_report = None if getattr(self, "expanded_report", None) == report_id else report_id
        self.refresh_reports()

    def _edit_report(self, report_id):
        row = self.db.get("reports", report_id)
        if row is None:
            return
        self.prepare_report(row)

    def _delete_report(self, report_id):
        row = self.db.get("reports", report_id)
        if row is None:
            return
        if not messagebox.askyesno(APP_NAME, "Bu rapor ve kayıtlı PDF silinsin mi?"):
            return
        path = Path(row["saved_path"] or "")
        self.db.delete("reports", report_id)
        if path.is_file() and not path.is_symlink() and path.suffix.lower() == ".pdf":
            try:
                path.unlink()
            except OSError:
                pass
        if getattr(self, "expanded_report", None) == report_id:
            self.expanded_report = None
        self.refresh_reports()

    def _open_report(self, path):
        file = Path(path or "")
        if not file.is_file() or file.is_symlink() or file.suffix.lower() != ".pdf":
            messagebox.showwarning(APP_NAME, "PDF bulunamadı.")
            return
        open_path(file)

    def _edit_invoice(self, row_id):
        row = self.db.get("invoices", row_id)
        if row is None:
            return
        dialog = RecordDialog(self, SCHEMAS["Faturalar"], row, self.watch_dir())
        self.wait_window(dialog)
        if dialog.result:
            self.db.save("invoices", dialog.result, row["id"])
            fresh = self.db.get("invoices", row["id"])
            if fresh is not None:
                total, paid = invoice_money(fresh)
                stored = float(fresh["paid_amount"] or 0) if fresh["paid_amount"] else 0
                if fresh["status"] == "Ödendi" and stored < total:
                    self.db.save("invoices", {"paid_amount": total}, fresh["id"])
                elif fresh["status"] == "Ödenmedi" and stored:
                    self.db.save("invoices", {"paid_amount": 0}, fresh["id"])
            self.selected_invoice = row["id"]
            self.refresh_current()

    def _phone_for_party(self, party):
        target = fold(party)
        if not target:
            return ""
        for contact in self.db.all("contacts"):
            names = [contact["name"] or ""]
            if "contact_person" in contact.keys():
                names.append(contact["contact_person"] or "")
            if any(fold(name) == target for name in names) and contact["phone"]:
                return contact["phone"]
        return ""

    def _log_invoice_event(self, row_id, kind, text):
        row = self.db.get("invoices", row_id)
        if row is None:
            return
        events = invoice_events(row)
        events.append({"kind": kind, "at": datetime.now().strftime("%Y-%m-%d %H:%M"), "text": text})
        self.db.save("invoices", {"events_json": json.dumps(events, ensure_ascii=False)}, row_id)

    def collect_invoice(self, row_id):
        row = self.db.get("invoices", row_id)
        if row is None:
            return
        if row["status"] == "İptal":
            messagebox.showinfo(APP_NAME, "İptal edilmiş faturada tahsilat kaydı açılmaz.")
            return
        total, paid = invoice_money(row)
        remaining = round(total - paid, 2)
        if total <= 0:
            messagebox.showinfo(APP_NAME, "Önce faturaya tutar yaz.")
            return
        if remaining <= 0:
            messagebox.showinfo(APP_NAME, "Bu fatura zaten tahsil edilmiş.")
            return
        answer = simpledialog.askstring(
            APP_NAME,
            f"Tahsil edilen tutar    kalan {format_amount(remaining)} ₺",
            initialvalue=format_amount(remaining),
            parent=self,
        )
        if answer is None:
            return
        try:
            amount = parse_amount(answer, True)
        except ValueError as exc:
            messagebox.showwarning(APP_NAME, str(exc))
            return
        if amount <= 0 or amount > remaining + 0.001:
            messagebox.showwarning(APP_NAME, f"Tutar 0 ile {format_amount(remaining)} ₺ arasında olmalı.")
            return
        new_paid = round(paid + amount, 2)
        status = "Ödendi" if new_paid >= total - 0.001 else "Kısmi ödendi"
        self.db.save("invoices", {"paid_amount": new_paid, "status": status}, row_id)
        self._log_invoice_event(row_id, "payment", f"{format_amount(amount)} ₺ tahsil edildi")
        self.selected_invoice = row_id
        self.refresh_invoices()

    def reach_invoice_party(self, row_id, channel):
        row = self.db.get("invoices", row_id)
        if row is None:
            return
        phone = self._phone_for_party(row["party"])
        digits = phone_digits(phone)
        if len(digits) < 11:
            messagebox.showinfo(APP_NAME, "Bu cari için rehberde telefon yok. Firmalar bölümünden ekleyebilirsin.")
            return
        total, paid = invoice_money(row)
        remaining = max(total - paid, 0)
        note = f"Merhaba, {row['invoice_no'] or 'fatura'} için kalan tutar {format_amount(remaining) or '0,00'} ₺."
        if row["due_date"]:
            note += f" Vade: {row['due_date']}."
        if channel == "whatsapp":
            webbrowser.open("https://wa.me/" + digits + "?text=" + urllib.parse.quote(note))
            self._log_invoice_event(row_id, "whatsapp", "WhatsApp yazışması açıldı")
        else:
            if sys.platform == "darwin":
                subprocess.Popen(["open", f"sms:{digits}"])
            else:
                webbrowser.open(f"sms:{digits}")
            self._log_invoice_event(row_id, "sms", "SMS taslağı açıldı")
        self.selected_invoice = row_id
        self.refresh_invoices()

    def cancel_invoice_record(self, row_id):
        row = self.db.get("invoices", row_id)
        if row is None or row["status"] == "İptal":
            return
        if not messagebox.askyesno(APP_NAME, "Bu fatura iptal edilsin mi? Kayıt silinmez, durumu İptal olur."):
            return
        self.db.save("invoices", {"status": "İptal"}, row_id)
        self._log_invoice_event(row_id, "cancel", "Fatura iptal edildi")
        self.selected_invoice = row_id
        self.refresh_invoices()

    def add_offer(self):
        if self.section != "Projeler":
            return
        selected = self.tree.selection()
        if not selected:
            messagebox.showinfo(APP_NAME, "Önce bir proje seç. Teklif, seçili işe eklenir.")
            return
        project = self.db.get("projects", int(selected[0]))
        if project is None:
            return
        self._edit_offer(project, None)

    def reanalyze_offers(self, project_id, focus_id, typed_pros, typed_cons, reread_focus):
        rows = list(self.db.offers_for(project_id))
        texts = {}
        for row in rows:
            cached = row["source_text"] if "source_text" in row.keys() else ""
            if row["id"] == focus_id and reread_focus:
                cached = ""
            url = (row["url"] if "url" in row.keys() else "") or ""
            if not cached and row["file_path"]:
                try:
                    cached = read_offer_text(row["file_path"])
                except (OSError, ValueError, subprocess.SubprocessError, zipfile.BadZipFile, ET.ParseError):
                    cached = ""
            if url and "Adres: " not in (cached or ""):
                cached = "\n".join(part for part in (cached, f"Adres: {offer_url_slug_text(url)}") if part)
            texts[row["id"]] = (cached or "")[:20000]
        profiles = []
        for row in rows:
            amount = float(row["amount"]) if row["amount"] is not None else None
            has_source = bool(row["file_path"] or ("url" in row.keys() and row["url"]))
            profiles.append(extract_offer_profile(row["vendor"], amount, texts[row["id"]], has_source))
        for row, profile in zip(rows, profiles):
            pros_lines, cons_lines = narrate_offer(profile, profiles)
            pros_typed = typed_pros if row["id"] == focus_id else (row["pros"] or "")
            cons_typed = typed_cons if row["id"] == focus_id else (row["cons"] or "")
            self.db.save("offers", {
                "pros": compose_offer_text(pros_typed, pros_lines),
                "cons": compose_offer_text(cons_typed, cons_lines),
                "source_text": texts[row["id"]],
            }, row["id"])

    def _edit_offer(self, project, row):
        previous = (row["file_path"] or "") if row else ""
        dialog = OfferDialog(self, project, row)
        self.wait_window(dialog)
        if not dialog.result:
            return
        if row:
            saved_id = self.db.save("offers", dialog.result, row["id"])
        else:
            saved_id = self.db.save("offers", dialog.result)
        stored = dialog.result.get("file_path") or ""
        if previous and previous != stored:
            remove_offer_file(previous)
        reread = bool(stored) and previous != stored and not dialog.result.get("source_text")
        self.status.configure(text="Teklif dosyaları okunuyor…")
        self.update_idletasks()
        self.reanalyze_offers(project["id"], saved_id, dialog.result.get("pros") or "", dialog.result.get("cons") or "", reread)
        if self.tree.exists(str(project["id"])):
            self.tree.selection_set(str(project["id"]))
        self.refresh_list()

    def _delete_offer(self, project_id, offer_id):
        if not messagebox.askyesno(APP_NAME, "Bu teklif silinsin mi?"):
            return
        row = self.db.get("offers", offer_id)
        if row:
            remove_offer_file(row["file_path"])
            self.db.delete("offers", offer_id)
        self.reanalyze_offers(project_id, None, "", "", False)
        if self.tree.exists(str(project_id)):
            self.tree.selection_set(str(project_id))
        self.refresh_list()

    def open_offer_compare(self, project):
        offers = list(self.db.offers_for(project["id"]))
        if len(offers) < 2:
            messagebox.showinfo(APP_NAME, "Karşılaştırma için bu projede en az iki teklif olmalı.")
            return
        project = self.db.get("projects", project["id"]) or project
        OfferCompareDialog(self, project, offers)

    def _open_offer_url(self, url):
        try:
            url = clean_offer_url(url)
        except ValueError as exc:
            messagebox.showwarning(APP_NAME, str(exc))
            return
        if url:
            webbrowser.open(url)

    def _open_offer_file(self, path):
        try:
            resolved = openable_offer(path)
        except ValueError as exc:
            messagebox.showwarning(APP_NAME, str(exc))
            return
        open_path(resolved)

    def refresh_offers(self):
        if not hasattr(self, "offer_panel"):
            return
        for child in self.offer_panel.winfo_children():
            child.destroy()
        if self.section != "Projeler":
            return
        selected = self.tree.selection()
        project = self.db.get("projects", int(selected[0])) if selected else None
        head = tk.Frame(self.offer_panel, bg=CANVAS)
        head.pack(fill="x")
        title = "Teklifler" if project is None else f"Teklifler  ·  {project['name']}"
        tk.Label(head, text=title, bg=CANVAS, fg=INK, font=(FONT, 14, "bold")).pack(side="left")
        if project is not None:
            Click(head, "Teklif ekle", self.add_offer, ACCENT, "white", ACCENT_HOVER, font=(FONT, 12, "bold"), padx=12, pady=6).pack(side="right")
            Click(head, "Karşılaştır", lambda current=project: self.open_offer_compare(current), NAV, "white", NAV_ON, font=(FONT, 12, "bold"), padx=12, pady=6).pack(side="right", padx=(0, 8))
        if project is None:
            tk.Label(self.offer_panel, text="Bir proje seç. Teklifler, avantaj ve dezavantajlarıyla burada karşılaştırılır.", bg=CANVAS, fg=MUTED, font=(FONT, 12), anchor="w").pack(anchor="w", pady=(6, 0))
            return
        offers = self.db.offers_for(project["id"])
        amounts = [float(row["amount"]) for row in offers if row["amount"] is not None]
        low = min(amounts) if amounts else None
        high = max(amounts) if amounts else None
        budget = float(project["budget"]) if project["budget"] is not None else None
        if not offers:
            tk.Label(self.offer_panel, text="Bu iş için henüz teklif yok. PDF, kayıtlı e-posta, görüntü, Word veya Excel ekleyebilirsin.", bg=CANVAS, fg=MUTED, font=(FONT, 12), anchor="w", wraplength=860, justify="left").pack(anchor="w", pady=(6, 0))
            return
        firms = list(dict.fromkeys(row["vendor"] or "Firma" for row in offers))
        parts = [f"{len(firms)} firma · {len(offers)} teklif"]
        if low is not None:
            best = [row["vendor"] for row in offers if row["amount"] is not None and float(row["amount"]) == low]
            parts.append(f"En avantajlı tutar: {', '.join(best)} · {format_amount(low)} ₺")
        if high is not None and low is not None and high != low:
            parts.append(f"En yüksek: {format_amount(high)} ₺")
            parts.append(f"Fark: {format_amount(high - low)} ₺")
        if budget is not None and low is not None:
            gap = budget - low
            if gap >= 0:
                parts.append(f"En düşük teklif bütçenin {format_amount(gap)} ₺ altında")
            else:
                parts.append(f"En düşük teklif bütçeyi {format_amount(-gap)} ₺ aşıyor")
        tk.Label(self.offer_panel, text="   ·   ".join(parts), bg=CANVAS, fg=MUTED, font=(FONT, 12), anchor="w", wraplength=900, justify="left").pack(anchor="w", pady=(4, 6))
        canvas = tk.Canvas(self.offer_panel, bg=CANVAS, highlightthickness=0, height=320)
        scroll = ttk.Scrollbar(self.offer_panel, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scroll.set)
        canvas.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        inner = tk.Frame(canvas, bg=CANVAS)
        window = canvas.create_window((0, 0), window=inner, anchor="nw")
        inner.bind("<Configure>", lambda _e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda event: canvas.itemconfigure(window, width=event.width))

        def wheel(event):
            canvas.yview_scroll(-1 if event.delta > 0 else 1, "units")

        def bind_wheel(widget):
            widget.bind("<MouseWheel>", wheel)
            for child in widget.winfo_children():
                bind_wheel(child)

        ordered = sorted(offers, key=lambda row: (fold(row["vendor"] or ""), row["amount"] is None, float(row["amount"] or 0)))
        seen = None
        for row in ordered:
            vendor = row["vendor"] or "Firma"
            key = fold(vendor)
            if key != seen:
                tk.Label(inner, text=vendor, bg=CANVAS, fg=ACCENT, font=(FONT, 13, "bold"), anchor="w").pack(fill="x", pady=(8, 2))
                seen = key
            self._offer_card(inner, project, row, low, high, budget)
        bind_wheel(inner)
        canvas.bind("<MouseWheel>", wheel)
        inner.update_idletasks()
        canvas.configure(scrollregion=canvas.bbox("all"))

    def _offer_card(self, parent, project, row, low, high, budget):
        amount = float(row["amount"]) if row["amount"] is not None else None
        cheapest = amount is not None and low is not None and high is not None and amount == low and low != high
        dearest = amount is not None and low is not None and high is not None and amount == high and low != high
        edge = ACCENT if cheapest else LINE
        card = tk.Frame(parent, bg=CARD, highlightbackground=edge, highlightthickness=1)
        card.pack(fill="x", pady=(0, 8))
        top = tk.Frame(card, bg=CARD)
        top.pack(fill="x", padx=12, pady=(10, 0))
        tk.Label(top, text=row["vendor"], bg=CARD, fg=INK, font=(FONT, 14, "bold"), anchor="w").pack(side="left")
        amount_text = f"{format_amount(amount)} ₺" if amount is not None else "Tutar yok"
        tk.Label(top, text=amount_text, bg=CARD, fg=OK if cheapest else INK, font=(FONT, 14, "bold")).pack(side="right")
        badges = tk.Frame(card, bg=CARD)
        badges.pack(fill="x", padx=12, pady=(4, 0))
        if cheapest:
            tk.Label(badges, text="En avantajlı tutar", bg=SOFT, fg=OK, font=(FONT, 10, "bold"), padx=6, pady=1).pack(side="left")
        elif dearest:
            tk.Label(badges, text="En yüksek tutar", bg="#fdecea", fg=DANGER, font=(FONT, 10, "bold"), padx=6, pady=1).pack(side="left")
        if amount is not None and low is not None and amount > low:
            tk.Label(badges, text=f"{format_amount(amount - low)} ₺ daha yüksek", bg="#fff6e6", fg=WARN, font=(FONT, 10, "bold"), padx=6, pady=1).pack(side="left", padx=(6, 0))
        if amount is not None and budget is not None:
            if amount <= budget:
                tk.Label(badges, text="Bütçe içinde", bg=SOFT, fg=OK, font=(FONT, 10, "bold"), padx=6, pady=1).pack(side="left", padx=(6, 0))
            else:
                tk.Label(badges, text=f"Bütçeyi {format_amount(amount - budget)} ₺ aşıyor", bg="#fdecea", fg=DANGER, font=(FONT, 10, "bold"), padx=6, pady=1).pack(side="left", padx=(6, 0))
        meta_bits = [row["offer_date"] or "tarih yok"]
        if row["title"]:
            meta_bits.append(row["title"])
        if row["file_path"]:
            meta_bits.append(offer_kind(row["file_path"]))
        page = (row["url"] if "url" in row.keys() else "") or ""
        if page:
            meta_bits.append(f"İnternet sayfası · {offer_url_host(page)}")
        tk.Label(card, text="  ·  ".join(meta_bits), bg=CARD, fg=MUTED, font=(FONT, 11), anchor="w").pack(fill="x", padx=12, pady=(4, 0))
        split = tk.Frame(card, bg=CARD)
        split.pack(fill="x", padx=12, pady=(8, 0))
        split.columnconfigure(0, weight=1)
        split.columnconfigure(1, weight=1)
        self._offer_side(split, 0, "Avantajlar", row["pros"], SOFT, OK)
        self._offer_side(split, 1, "Dezavantajlar", row["cons"], "#fdecea", DANGER)
        if row["notes"]:
            tk.Label(card, text=one_line(row["notes"], 180), bg=CARD, fg=MUTED, font=(FONT, 11), anchor="w", wraplength=760, justify="left").pack(fill="x", padx=12, pady=(6, 0))
        actions = tk.Frame(card, bg=CARD)
        actions.pack(fill="x", padx=12, pady=(8, 10))
        if row["file_path"]:
            Click(actions, "Dosyayı aç", lambda path=row["file_path"]: self._open_offer_file(path), SOFT, INK, SOFT_HOVER, font=(FONT, 11), padx=10, pady=4).pack(side="left", padx=(0, 6))
        if page:
            Click(actions, "Sayfayı aç", lambda url=page: self._open_offer_url(url), SOFT, INK, SOFT_HOVER, font=(FONT, 11), padx=10, pady=4).pack(side="left", padx=(0, 6))
        Click(actions, "Düzenle", lambda current=row: self._edit_offer(project, current), CARD, INK, HOVER, font=(FONT, 11), padx=10, pady=4, border=LINE).pack(side="left", padx=(0, 6))
        Click(actions, "Sil", lambda offer_id=row["id"]: self._delete_offer(project["id"], offer_id), CARD, DANGER, "#fdecea", font=(FONT, 11), padx=10, pady=4, border=LINE).pack(side="left")

    def _offer_side(self, parent, column, title, text, background, color):
        box = tk.Frame(parent, bg=background)
        box.grid(row=0, column=column, sticky="nsew", padx=(0 if column == 0 else 6, 0))
        tk.Label(box, text=title, bg=background, fg=color, font=(FONT, 11, "bold"), anchor="w").pack(fill="x", padx=8, pady=(6, 0))
        lines = [line.strip(" •-\t") for line in str(text or "").splitlines() if line.strip()]
        body = "\n".join(f"• {line}" for line in lines) if lines else "Yazılmadı"
        tk.Label(box, text=body, bg=background, fg=INK if lines else MUTED, font=(FONT, 11), anchor="w", justify="left", wraplength=340).pack(fill="x", padx=8, pady=(2, 8))

    def delete_record(self):
        if self.section == "Raporlar":
            return
        if self.section == "Firmalar" and self.tree.selection() and self._address_hit_for(self.tree.selection()[0]):
            messagebox.showinfo(APP_NAME, "Adres defterindeki kişi buradan silinmez. Kalemle rehbere ekleyebilirsin.")
            return
        row = self.selected_row()
        if row is None:
            return
        config = self.schema()
        linked = []
        others = []
        if config["table"] == "projects":
            linked = [task for task in self.db.all("tasks") if "project_id" in task.keys() and task["project_id"] == row["id"]]
            others = [project for project in self.db.all("projects") if project["id"] != row["id"]]
            if linked and not others:
                messagebox.showinfo(APP_NAME, "Bu projedeki kayıtlar duruyor. Başka bir proje kalmadan bu proje silinemez.")
                return
        question = f"Bu {config['title'].lower()} kaydı silinsin mi?"
        if linked:
            question += f"\n{len(linked)} kayıt başka projeye taşınır. Teklifler silinir."
        if not messagebox.askyesno(APP_NAME, question):
            return
        if config["table"] == "projects" and linked:
            target = next((project for project in others if (project["proj_key"] or "").upper() == "GEN"), others[0])
            for task in linked:
                self.db.save("tasks", {"project_id": target["id"], "issue_no": self.db.next_issue_no(target["id"])}, task["id"])
            if self.board_project_id == row["id"]:
                self.board_project_id = target["id"]
        if config["table"] == "projects":
            self.db.delete_project_offers(row["id"])
        if config["table"] == "tasks":
            remove_note_dir(row["id"])
        if config["table"] == "transactions":
            raw = row["docs_json"] if "docs_json" in row.keys() else ""
            for item in ledger_documents(raw):
                remove_ledger_file(item["path"])
        if config["table"] == "invoices" and row["source_key"]:
            stamp = ""
            if row["attachment"]:
                try:
                    stamp = str(Path(row["attachment"]).stat().st_mtime_ns)
                except OSError:
                    stamp = ""
            self.db.ignore(f"{row['source_key']}|{stamp}" if row["source"] == "klasor" else row["source_key"])
        self.db.delete(config["table"], row["id"])
        if self.selected_invoice == row["id"]:
            self.selected_invoice = None
        self.refresh_current()

    def export_excel(self):
        if self.section in ("İş takibi", "Raporlar"):
            messagebox.showinfo(APP_NAME, "Excel için Faturalar, Görevler, Rehber, Gelir-Gider veya Projeler bölümünü aç.")
            return
        sheets = self._excel_sheets()
        title = PAGE_TITLES.get(self.section, self.section)
        path = filedialog.asksaveasfilename(
            parent=self,
            title="Excel dışa aktar",
            defaultextension=".xlsx",
            initialfile=f"{title}.xlsx",
            filetypes=[("Excel", "*.xlsx")],
        )
        if not path:
            return
        if not str(path).lower().endswith(".xlsx"):
            path = f"{path}.xlsx"
        try:
            write_workbook(path, sheets)
        except (OSError, ValueError, ET.ParseError) as exc:
            messagebox.showwarning(APP_NAME, f"Excel yazılamadı.\n{exc}")
            return
        count = sum(len(sheet["rows"]) for sheet in sheets)
        self.status.configure(text=f"Excel kaydedildi: {Path(path).name}  ·  {count} satır")

    def _excel_sheets(self):
        if self.section == "Faturalar":
            rows = self.current_invoices()
            return [_export_sheet("Faturalar", _INVOICE_EXPORT, rows)]
        if self.section == "Görevler":
            columns = [
                ("Anahtar", "text", lambda row: self.issue_key(row)) if column[0] == "Anahtar" else column
                for column in _TASK_EXPORT
            ]
            return [_export_sheet("Görevler", columns, self.visible_rows())]
        if self.section == "Hatırlatmalar":
            return [_export_sheet("Hatırlatmalar", _REMINDER_EXPORT, self.visible_rows())]
        if self.section == "Firmalar":
            return [_export_sheet("Rehber", _CONTACT_EXPORT, self.visible_rows())]
        if self.section == "Gelir-Gider":
            return [_export_sheet("Gelir-Gider", _LEDGER_EXPORT, self.visible_rows())]
        projects = self.visible_rows()
        counts = {project["id"]: self.db.offer_count(project["id"]) for project in projects}
        project_columns = [
            column if column[0] != "Teklif sayısı" else ("Teklif sayısı", "text", lambda row, counts=counts: counts.get(row["id"], 0))
            for column in _PROJECT_EXPORT
        ]
        offers = []
        for project in projects:
            for offer in self.db.offers_for(project["id"]):
                offers.append((project, offer))
        return [
            _export_sheet("Projeler", project_columns, projects),
            _export_sheet("Teklifler", _OFFER_EXPORT, offers),
        ]

    def backup(self):
        path = filedialog.asksaveasfilename(parent=self, title="Veritabanını yedekle", defaultextension=".db", initialfile=f"is_merkezi-{today()}.db", filetypes=[("SQLite", "*.db")])
        if not path:
            return
        self.db.backup_to(path)
        messagebox.showinfo(APP_NAME, f"Yedek alındı:\n{path}")

    def open_settings(self):
        values = {
            "watch_dir": str(self.watch_dir()),
            "excel_path": str(self.excel_path()),
            "mail_enabled": self.db.get_setting("mail_enabled", "0"),
            "mail_host": self.db.get_setting("mail_host", ""),
            "mail_port": self.db.get_setting("mail_port", "993"),
            "mail_user": self.db.get_setting("mail_user", ""),
            "mail_box": self.db.get_setting("mail_box", "INBOX"),
        }
        dialog = SettingsDialog(self, values)
        self.wait_window(dialog)
        if not dialog.result:
            return
        result = dialog.result
        password = result.pop("password")
        read_now = result.pop("read_now", False)
        for key, value in result.items():
            self.db.set_setting(key, value)
        if password:
            try:
                keychain_set(result["mail_user"], password)
            except ValueError as exc:
                messagebox.showwarning(APP_NAME, str(exc))
                return
        self._refresh_mail_label()
        self.sync_now()
        if read_now:
            self._start_mail()

    def sync_now(self):
        try:
            note = self.sync_folder() or "Klasör güncel."
        except (OSError, ValueError, zipfile.BadZipFile) as exc:
            note = f"Klasör okunamadı: {exc}"
        self.status.configure(text=note)
        self.refresh_current()

    def sync_folder(self):
        watch = self.watch_dir()
        if not acceptable_watch_dir(watch):
            return "Fatura klasörü ayarlı değil."
        excel = self.excel_path()
        if excel.exists() and (excel.is_symlink() or excel.suffix.lower() != ".xlsx" or not is_within(excel, watch)):
            return "Excel dosyası fatura klasörünün dışında, yazılmadı."
        files = list_invoice_files(watch)
        known = {}
        if excel.is_file():
            for row in read_excel_invoices(excel):
                known[path_key(row["path"])] = row
        added = 0
        pending = []
        for path in files:
            key = str(path.resolve())
            try:
                stamp = str(path.stat().st_mtime_ns)
            except OSError:
                continue
            if self.db.ignored(f"{key}|{stamp}") or self.db.find_source(key):
                continue
            meta = known.get(key, {})
            amount = meta.get("total")
            self.db.save(
                "invoices",
                {
                    "kind": "Gelen",
                    "party": (meta.get("seller") or guess_party(path, watch) or "Firma")[:160],
                    "invoice_no": (meta.get("doc_no") or meta.get("folder") or path.parent.name)[:80],
                    "description": path.name,
                    "amount": amount,
                    "invoice_date": meta.get("invoice_date") or "",
                    "due_date": meta.get("invoice_date") or "",
                    "status": "Ödenmedi",
                    "attachment": key,
                    "source": "klasor",
                    "source_key": key,
                    "currency": meta.get("currency") or "TRY",
                    "tax_no": meta.get("tax_no") or "",
                    "excel_status": meta.get("excel_status") or "Kontrol Gerekli",
                    "file_name": path.name,
                },
            )
            added += 1
            if key not in known:
                pending.append(
                    (
                        path.parent.name if path.parent.resolve() != watch.resolve() else path.stem,
                        key,
                        path.name,
                        "",
                        date.today(),
                        guess_party(path, watch),
                        "",
                        None,
                        None,
                        "TRY",
                        None,
                        "Kontrol Gerekli",
                        date.today(),
                    )
                )
        written = 0
        if pending:
            written = append_excel_invoices(excel, pending)
        if not added and not written:
            return ""
        return f"{added} fatura uygulamaya alındı, {written} satır Excel'e yazıldı."

    def sync_mail_now(self):
        if self.db.get_setting("mail_enabled") != "1":
            messagebox.showinfo(APP_NAME, "E-posta takibi kapalı. Ayarlardan açabilirsin.")
            return
        self._start_mail()

    def _start_mail(self, allow_launch=True):
        if self._mail_busy:
            return
        local = self.db.get_setting("mail_host", "") == LOCAL_MAIL_HOST
        user = self.db.get_setting("mail_user", "")
        password = "" if local else keychain_get(user)
        if not local and not password:
            self.status.configure(text="E-posta parolası Anahtar Zinciri'nde yok.")
            return
        self._mail_busy = True
        settings = {
            "local": local,
            "allow_launch": allow_launch,
            "host": self.db.get_setting("mail_host", ""),
            "port": int(self.db.get_setting("mail_port", "993") or 993),
            "user": user,
            "password": password,
            "mailbox": self.db.get_setting("mail_box", "INBOX") or "INBOX",
        }
        threading.Thread(target=self._mail_worker, args=(settings,), daemon=True).start()

    def _mail_worker(self, settings):
        try:
            if settings.get("local"):
                found = fetch_mail_app_invoices(settings.get("allow_launch", True))
                if found is None:
                    self.events.put(("mail", [], "", "skip"))
                    return
                self.events.put(("mail", found, "", "auto" if not settings.get("allow_launch", True) else ""))
                return
            found = fetch_mail_invoices(settings["host"], settings["port"], settings["user"], settings["password"], settings["mailbox"])
            self.events.put(("mail", found, "", "auto" if not settings.get("allow_launch", True) else ""))
        except Exception as exc:
            mode = "auto" if not settings.get("allow_launch", True) else ""
            self.events.put(("mail", [], one_line(exc, 180), mode))

    def _apply_mail(self, found, error, mode=""):
        self._mail_busy = False
        if mode == "skip":
            return
        if error:
            if mode != "auto":
                self.status.configure(text=f"E-posta alınamadı: {error}")
            return
        added = 0
        pending = []
        for item in found:
            if self.db.find_source(item["source_key"]) or self.db.ignored(item["source_key"]):
                continue
            self.db.save(
                "invoices",
                {
                    "kind": "Gelen",
                    "party": item["party"],
                    "invoice_no": "",
                    "description": item["description"],
                    "amount": None,
                    "invoice_date": today(),
                    "due_date": "",
                    "status": "Ödenmedi",
                    "attachment": item["attachment"],
                    "source": "eposta",
                    "source_key": item["source_key"],
                    "currency": "TRY",
                    "excel_status": "E-postadan alındı",
                    "file_name": item["filename"],
                },
            )
            added += 1
            pending.append((item["filename"], item["attachment"], item["filename"], "", date.today(), item["party"], "", None, None, "TRY", None, "E-postadan alındı", date.today()))
        written = 0
        if pending:
            try:
                written = append_excel_invoices(self.excel_path(), pending)
            except (OSError, ValueError, zipfile.BadZipFile) as exc:
                self.status.configure(text=f"E-posta ekleri uygulamada. Excel yazılamadı: {exc}")
                self.refresh_current()
                return
        if added == 0 and mode == "auto":
            return
        self.status.configure(text=f"E-posta kontrol edildi. {added} yeni fatura, {written} Excel satırı.")
        self.refresh_current()

    def tick_folder(self):
        if not self._alive:
            return
        try:
            note = self.sync_folder()
            if note:
                self.status.configure(text=note)
                self.refresh_current()
        except (OSError, ValueError, zipfile.BadZipFile):
            pass
        if self.db.get_setting("mail_enabled") == "1":
            self._start_mail(False)
        self.after(20000, self.tick_folder)

    def _drain(self):
        if not self._alive:
            return
        try:
            while True:
                event = self.events.get_nowait()
                kind = event[0]
                if kind == "preview":
                    _, pane_id, generation, image = event
                    pane = self.board_preview if id(self.board_preview) == pane_id else self.invoice_preview
                    pane.receive(generation, image)
                elif kind == "lightbox":
                    _, pane_id, _generation, image, title = event
                    pane = self.board_preview if id(self.board_preview) == pane_id else self.invoice_preview
                    pane.open_lightbox(image, title)
                elif kind == "mail":
                    self._apply_mail(event[1], event[2], event[3] if len(event) > 3 else "")
        except queue.Empty:
            pass
        self.after(150, self._drain)

    def check_due(self):
        if not self._alive:
            return
        self._scan_due()
        self.after(5000, self.check_due)

    def _scan_due(self):
        if not self._alive:
            return
        try:
            self._check_due()
        except sqlite3.Error:
            pass

    def show_toast(self, title, body, when="", badge="HATIRLATMA", actions=None, sla_id=None):
        try:
            self._open_toast(title, body, when, badge, actions or [], sla_id)
        except tk.TclError:
            return

    def _open_toast(self, title, body, when, badge="HATIRLATMA", actions=None, sla_id=None):
        toast = tk.Toplevel(self)
        toast.title("Hatırlatma")
        toast.configure(bg=CARD)
        toast.resizable(False, False)
        toast._sla_id = sla_id
        try:
            toast.attributes("-topmost", True)
        except tk.TclError:
            pass
        shell = tk.Frame(toast, bg=CARD, highlightbackground="#8aa0b5", highlightthickness=1)
        shell.pack(fill="both", expand=True)
        accent = WARN if sla_id else ACCENT
        tk.Frame(shell, bg=accent, height=4).pack(fill="x")
        head = tk.Frame(shell, bg=CARD)
        head.pack(fill="x", padx=14, pady=(10, 0))
        tk.Label(head, text=badge, bg=CARD, fg=accent, font=(FONT, 11, "bold")).pack(side="left")
        Click(head, "✕", lambda target=toast: self._close_toast(target), CARD, INK, HOVER, font=(FONT, 13, "bold"), padx=8, pady=2, border=LINE).pack(side="right")
        if when:
            tk.Label(shell, text=when, bg=CARD, fg=INK, font=(FONT, 13, "bold"), anchor="w", justify="left").pack(fill="x", padx=14, pady=(8, 0))
        tk.Label(shell, text=one_line(title, 80) or "Hatırlatma", bg=CARD, fg=INK, font=(FONT, 15, "bold"), anchor="w", justify="left", wraplength=340).pack(fill="x", padx=14, pady=(4, 0))
        if body:
            tk.Label(shell, text=one_line(body, 220), bg=CARD, fg=MUTED, font=(FONT, 12), anchor="w", justify="left", wraplength=340).pack(fill="x", padx=14, pady=(4, 0))
        if actions:
            bar = tk.Frame(shell, bg=CARD)
            bar.pack(fill="x", padx=14, pady=(10, 0))
            tk.Label(bar, text="Ertele", bg=CARD, fg=MUTED, font=(FONT, 11)).pack(side="left", padx=(0, 6))
            for label, command in actions:
                Click(bar, label, command, CARD, INK, HOVER, font=(FONT, 11, "bold"), padx=8, pady=4, border=LINE).pack(side="left", padx=(0, 6))
        tk.Frame(shell, bg=CARD, height=12).pack(fill="x")
        toast.protocol("WM_DELETE_WINDOW", lambda target=toast: self._close_toast(target))
        self._toasts = getattr(self, "_toasts", [])
        self._toasts.append(toast)
        if sla_id is not None:
            self._sla_open = getattr(self, "_sla_open", set())
            self._sla_open.add(sla_id)
        toast._toast_job = toast.after(120000, lambda target=toast: self._close_toast(target))
        toast.update_idletasks()
        toast._toast_h = max(toast.winfo_reqheight(), 96)
        self._place_toasts()
        toast.lift()

    def _close_toast(self, toast):
        job = getattr(toast, "_toast_job", None)
        if job is not None:
            try:
                toast.after_cancel(job)
            except tk.TclError:
                pass
            toast._toast_job = None
        sla_id = getattr(toast, "_sla_id", None)
        if sla_id is not None:
            self._sla_open = getattr(self, "_sla_open", set())
            self._sla_open.discard(sla_id)
        if toast.winfo_exists():
            toast.destroy()
        self._toasts = [item for item in getattr(self, "_toasts", []) if item is not toast and item.winfo_exists()]
        self._place_toasts()

    def _place_toasts(self):
        width, gap, stack = 440, 12, 0
        alive = [item for item in getattr(self, "_toasts", []) if item.winfo_exists()]
        self._toasts = alive
        screen_w = max(self.winfo_screenwidth(), width + 40)
        screen_h = max(self.winfo_screenheight(), 200)
        for toast in reversed(alive):
            height = getattr(toast, "_toast_h", 120)
            x = max(16, screen_w - width - 24)
            y = max(48, screen_h - height - 96 - stack)
            toast.geometry(f"{width}x{height}+{x}+{y}")
            stack += height + gap
            try:
                toast.attributes("-topmost", True)
            except tk.TclError:
                pass
            toast.deiconify()
            toast.lift()

    def _check_due(self):
        today_date = date.today()
        sent = 0
        for row in self.db.all("reminders"):
            if sent >= 5 or row["status"] != "Aktif":
                continue
            if not due_on_schedule(row["remind_date"], row["repeat"], today_date) or not time_reached(row["remind_time"]):
                continue
            clock = row["remind_time"] or ""
            if (row["repeat"] or "Yok") == "Yok":
                key = f"reminder:{row['id']}:{row['remind_date']}:{clock}"
            else:
                key = f"reminder:{row['id']}:{today()}:{clock}"
            if self.db.seen(key):
                continue
            day = parse_day(row["remind_date"])
            when = f"{day.day} {MONTHS[day.month - 1]} {day.year}" if day else (row["remind_date"] or "")
            if row["remind_time"]:
                when = f"{when} · {row['remind_time']}" if when else row["remind_time"]
            self.show_toast(row["title"], row["description"] or "", when)
            if sent == 0:
                play_alert()
            notify("Hatırlatma", row["title"])
            self.db.mark_seen(key)
            sent += 1
        for row in self.db.all("tasks"):
            if sent >= 5 or row["status"] in ("Tamamlandı", "İptal"):
                continue
            due = parse_day(row["due_date"])
            if due is None or due > today_date:
                continue
            key = f"task:{row['id']}:{row['due_date']}"
            if self.db.seen(key):
                continue
            notify("Görev son tarihi", row["title"])
            self.db.mark_seen(key)
            sent += 1
        for row in self.db.all("invoices"):
            if sent >= 5 or row["status"] not in ("Ödenmedi", "Kısmi ödendi"):
                continue
            due = parse_day(row["due_date"])
            if due is None or due > today_date:
                continue
            key = f"invoice:{row['id']}:{row['due_date']}"
            if self.db.seen(key):
                continue
            notify("Fatura vadesi", f"{row['party']} · {format_amount(row['amount'])} ₺")
            self.db.mark_seen(key)
            sent += 1
        self._check_waiting_sla()

    def _check_waiting_sla(self):
        now = datetime.now().replace(microsecond=0)
        open_ids = getattr(self, "_sla_open", set())
        shown = 0
        sounded = False
        refresh = False
        for row in self.db.all("tasks"):
            if row["status"] in ("Tamamlandı", "İptal") or row["id"] in open_ids:
                continue
            start = task_waiting_since(row)
            if start is None or now < start + timedelta(hours=24):
                continue
            moment = task_sla_moment(row)
            if moment is None or now < moment:
                continue
            if shown < 3:
                self._show_sla(row)
                self.db.save("tasks", {"sla_next": (now + timedelta(hours=2)).strftime("%Y-%m-%d %H:%M:%S")}, row["id"])
                if not sounded:
                    play_alert()
                    sounded = True
                shown += 1
            else:
                self.db.save("tasks", {"sla_next": (now + timedelta(minutes=2)).strftime("%Y-%m-%d %H:%M:%S")}, row["id"])
            refresh = True
        if refresh and self.section == "İş takibi":
            self.refresh_board()

    def _show_sla(self, row):
        start = task_waiting_since(row)
        waited = ""
        if start is not None:
            hours = max(24, int((datetime.now() - start).total_seconds() // 3600))
            waited = f"{hours} saattir bekliyor"
        who = row["assignee"] or "sahipsiz"
        task_id = row["id"]
        actions = [(f"{hours} saat", lambda span=hours, current=task_id: self._snooze_sla(current, span)) for hours in (2, 4, 8)]
        self.show_toast(
            row["title"],
            "Bekliyor listesinde 24 saat doldu. İş tamamlanana kadar 2 saatte bir uyarır.",
            " · ".join(part for part in (waited, who) if part),
            badge="SLA",
            actions=actions,
            sla_id=task_id,
        )
        notify("Bekleyen iş", row["title"])

    def _snooze_sla(self, task_id, hours):
        when = datetime.now().replace(microsecond=0) + timedelta(hours=hours)
        self.db.save("tasks", {"sla_next": when.strftime("%Y-%m-%d %H:%M:%S")}, task_id)
        for toast in list(getattr(self, "_toasts", [])):
            if getattr(toast, "_sla_id", None) == task_id:
                self._close_toast(toast)
        if self.section == "İş takibi":
            self.refresh_board()

    def on_close(self):
        self._alive = False
        try:
            self.db.close()
        except sqlite3.Error:
            pass
        self.destroy()


def main():
    if sys.platform == "win32":
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except (AttributeError, OSError):
            pass
    try:
        app = App()
    except Exception:
        import traceback
        text = traceback.format_exc()
        try:
            DATA_DIR.mkdir(parents=True, exist_ok=True)
            (DATA_DIR / "baslangic.log").write_text(text, encoding="utf-8")
        except OSError:
            pass
        if sys.platform == "win32":
            ctypes.windll.user32.MessageBoxW(None, ("Uygulama açılamadı.\n\n" + text)[-1800:], APP_NAME, 0x10)
            return
        raise
    if not app._locked_out:
        app.mainloop()


if __name__ == "__main__":
    main()
