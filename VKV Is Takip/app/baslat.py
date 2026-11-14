# -*- coding: utf-8 -*-
"""Windows açılışı. Hata olursa pencere gösterir ve Veriler/hata.log dosyasına yazar."""
import ctypes
import traceback
from pathlib import Path

import runpy

APP = Path(__file__).resolve().parent / "is_merkezi.py"
LOG = Path(__file__).resolve().parent.parent / "Veriler" / "hata.log"


def fail(text):
    try:
        LOG.parent.mkdir(parents=True, exist_ok=True)
        LOG.write_text(text, encoding="utf-8")
    except OSError:
        pass
    ctypes.windll.user32.MessageBoxW(None, ("Uygulama açılamadı.\n\n" + text)[-1800:], "VKV İş Takip", 0x10)


try:
    runpy.run_path(str(APP), run_name="__main__")
except SystemExit:
    raise
except Exception:
    fail(traceback.format_exc())
