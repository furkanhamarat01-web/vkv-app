"""Tek pencerenin kullanıcı, rol ve parola istekleri."""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent / "talep"))
from talep_client import ApiError, Store

__all__ = ["ApiError", "client"]

_store = None


def client():
    global _store
    if _store is None:
        _store = Store()
    return _store
