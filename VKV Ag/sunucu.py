# -*- coding: utf-8 -*-
"""Masaüstü kısayolu. Sunucu kodu vkv-backend projesindedir."""
import runpy
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1] / "vkv-backend"
if not (ROOT / "vkv_backend").is_dir():
    sys.exit("vkv-backend klasörü yok. Sunucuya o projeyi kopyala.")
sys.path.insert(0, str(ROOT))
runpy.run_module("vkv_backend", run_name="__main__")
