#!/usr/bin/env python3
"""Masaüstü kısayolu. Asıl pencere is-merkezi/ui/talep içindedir."""
import runpy
from pathlib import Path

runpy.run_path(
    str(Path(__file__).resolve().parents[1] / "is-merkezi" / "ui" / "talep" / "talep.py"),
    run_name="__main__",
)
