# app/layout_loader.py
from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from app.resource_path import resource_path


@dataclass(frozen=True)
class Layout:
    mapping: Dict[int, Tuple[int, int]]          # loom_no -> (row, col)
    dividers_v: List[int]                        # dikey ayraç çizgileri için kolon indexleri
    dividers_h: List[int]                        # yatay ayraç çizgileri için row indexleri
    max_row: int
    max_col: int


def _infer_bounds(mapping: Dict[int, Tuple[int, int]]) -> Tuple[int, int]:
    if not mapping:
        return (0, 0)
    max_r = max(rc[0] for rc in mapping.values())
    max_c = max(rc[1] for rc in mapping.values())
    return max_r, max_c


def load_layout_from_json(json_path: Path) -> Layout:
    data = json.loads(json_path.read_text(encoding="utf-8"))

    mapping: Dict[int, Tuple[int, int]] = {}

    # Tip A: "cells": [{"loom":1301,"row":0,"col":0}, ...]
    if "cells" in data:
        for cell in data["cells"]:
            loom = int(cell["loom"])
            r = int(cell["row"])
            c = int(cell["col"])
            mapping[loom] = (r, c)

    # Tip B: "mapping": {"1301":[0,0], ...}
    elif "mapping" in data:
        for k, v in data["mapping"].items():
            loom = int(k)
            r = int(v[0])
            c = int(v[1])
            mapping[loom] = (r, c)
    else:
        raise ValueError(f"Unsupported layout schema in {json_path.name}")

    div_v = [int(x) for x in data.get("dividers_v", [])]
    div_h = [int(x) for x in data.get("dividers_h", [])]

    mr, mc = _infer_bounds(mapping)
    return Layout(mapping=mapping, dividers_v=div_v, dividers_h=div_h, max_row=mr, max_col=mc)


def load_layout_for_site(site_key: str) -> Optional[Layout]:
    """
    site_key: ISKO14 | ISKO11 | MEKIKLI
    JSON yoksa None döner (hard-coded fallback için).
    """
    key = site_key.upper()
    rel = f"app/layouts/{key.lower()}.json"  # isko11.json, mekikli.json, isko14.json
    path = resource_path(rel)

    if not path.exists():
        return None

    return load_layout_from_json(path)
