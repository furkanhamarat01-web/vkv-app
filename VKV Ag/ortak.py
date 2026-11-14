# -*- coding: utf-8 -*-
"""VKV Ağ — sunucu ve istemcinin ortak sabitleri.

Bu klasör VKV İş Merkezi ve VKV Talep'ten ayrıdır.
Sunucu verisi sunucudadır. Her bilgisayarın istemci kaydı kendi kullanıcısındadır.
"""
import hashlib
import os
import platform
import secrets
import socket
import subprocess
import sys
from pathlib import Path

PORT = 8780
PUBLIC_SERVER = "http://hqb6bmzszcsc4t470l10npd5.185.184.210.52.sslip.io"
APP = "VKV Ağ"
FONT = ".AppleSystemUIFont" if sys.platform == "darwin" else "Segoe UI"
MAX_FILE = 20 * 1024 * 1024
FILE_EXT = {".pdf", ".png", ".jpg", ".jpeg", ".webp", ".txt", ".csv", ".xls", ".xlsx", ".doc", ".docx"}
ROLES = ("Yönetici", "Kullanıcı")
STATUSES = ("Açık", "Üzerinde", "Tamamlandı")
ONLINE_SECONDS = 20


def app_support():
    if sys.platform == "darwin":
        root = Path.home() / "Library" / "Application Support" / "VKVAg"
    elif sys.platform == "win32":
        root = Path(os.environ.get("APPDATA", str(Path.home()))) / "VKVAg"
    else:
        root = Path.home() / ".vkv_ag"
    root.mkdir(parents=True, exist_ok=True)
    return root


def sunucu_dir():
    chosen = os.environ.get("VKV_AG_HOME", "").strip().strip('"')
    root = Path(chosen) if chosen else app_support() / "sunucu"
    root.mkdir(parents=True, exist_ok=True)
    (root / "dosyalar").mkdir(parents=True, exist_ok=True)
    return root


def istemci_dir():
    root = app_support() / "istemci"
    root.mkdir(parents=True, exist_ok=True)
    return root


def hash_password(password, salt=None):
    salt = salt or secrets.token_bytes(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, 120_000)
    return salt.hex() + ":" + digest.hex()


def check_password(password, stored):
    try:
        salt_hex, digest_hex = str(stored).split(":", 1)
        digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), bytes.fromhex(salt_hex), 120_000)
    except (ValueError, TypeError):
        return False
    return secrets.compare_digest(digest.hex(), digest_hex)


def clean_username(value):
    return "".join(ch for ch in str(value or "").casefold() if ch.isalnum() or ch in "._-")[:32]


def safe_filename(name):
    base = Path(str(name or "")).name
    suffix = Path(base).suffix.lower()
    if suffix not in FILE_EXT:
        raise ValueError("PDF, görüntü, metin veya ofis dosyası paylaşılır.")
    stem = "".join(ch for ch in Path(base).stem if ch.isalnum() or ch in "._- ")[:80].strip() or "dosya"
    return stem + suffix


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
        "cihaz": (host or "Bu bilgisayar")[:80],
        "sistem": os_text[:80],
        "ip": ip[:64],
    }


def lan_ips():
    found = []
    try:
        infos = socket.getaddrinfo(socket.gethostname(), None, socket.AF_INET)
    except OSError:
        infos = []
    for info in infos:
        address = info[4][0]
        if address.startswith("127."):
            continue
        if address not in found:
            found.append(address)
    probe = device_snapshot()["ip"]
    if probe and not probe.startswith("127.") and probe not in found:
        found.insert(0, probe)
    return found


def api_base(value):
    raw = str(value or "").strip().rstrip("/")
    if not raw:
        raise ValueError("Sunucu adresi yaz.")
    if "://" in raw:
        scheme, rest = raw.split("://", 1)
        if scheme not in ("http", "https") or not rest or any(ch.isspace() for ch in rest):
            raise ValueError("Sunucu adresi geçersiz.")
        return f"{scheme}://{rest}"
    host, port = parse_address(raw)
    return f"http://{host}:{port}"


def parse_address(value):
    raw = str(value or "").strip()
    if "://" in raw:
        raw = raw.split("://", 1)[1]
    raw = raw.split("/", 1)[0].strip()
    if not raw:
        raise ValueError("Sunucu adresi yaz.")
    if raw.count(":") == 1:
        host, port_text = raw.rsplit(":", 1)
        if not port_text.isdigit():
            raise ValueError("Kapı numarası sayı olmalı.")
        port = int(port_text)
    else:
        host, port = raw, PORT
    host = host.strip()
    if not host or port < 1 or port > 65535:
        raise ValueError("Sunucu adresi geçersiz.")
    return host, port


def file_size(size):
    size = int(size or 0)
    if size < 1024:
        return f"{size} B"
    if size < 1024 * 1024:
        return f"{size / 1024:.0f} KB"
    return f"{size / (1024 * 1024):.1f} MB"
