from __future__ import annotations
import json
from pathlib import Path

OUT = Path("app/layouts/mekikli.json")
OUT.parent.mkdir(parents=True, exist_ok=True)

cells = []

# -----------------------------
# ALT BLOK (601–648)
# 6 satır x 8 sütun
# -----------------------------
start = 601
cols = 8
rows = 6

loom = start
for r in range(rows):
    row_index = (rows - 1) - r  # alttan yukarı
    for c in range(cols):
        cells.append({
            "loom": loom,
            "row": row_index,
            "col": c
        })
        loom += 1

# -----------------------------
# ÜST BLOK (649–700)
# 6 satır x 8 sütun
# -----------------------------
start = 649
loom = start
row_offset = rows + 2  # aradaki koridor boşluğu

for r in range(rows):
    row_index = row_offset + (rows - 1) - r
    for c in range(cols):
        if loom > 700:
            break
        cells.append({
            "loom": loom,
            "row": row_index,
            "col": c
        })
        loom += 1

data = {
    "site": "MEKIKLI",
    "schema": "cells-v1",
    "global_range": [601, 700],
    "cells": cells,
    "dividers_h": [rows],  # iki blok arası yatay koridor
    "dividers_v": [],
    "notes": "MEKIKLI birebir saha yerleşimi (2 blok, 8x6)"
}

OUT.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
print(f"Wrote {OUT} with {len(cells)} cells")
