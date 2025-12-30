# tools/gen_layout_isko11.py
from __future__ import annotations

import json
from pathlib import Path

# -------------------------------------------------------------------
# ISKO11: 6 blok (3 satır x 2 kolon)
# Numaralandırma: row-major (soldan sağa, sonra alt satır)
# Global aralık: 1301-1912 (612 adet)
# -------------------------------------------------------------------

OUT = Path("app/layouts/isko11.json")
OUT.parent.mkdir(parents=True, exist_ok=True)

START_LOOM = 1301
END_LOOM = 1912

# Blok düzeni
BLOCK_ROWS = 3
BLOCK_COLS = 2

# Her bloğun iç grid'i (fotoğrafa göre)
# Eğer sahada 1-2 satır/sütun fark edersen burayı düzeltip yeniden üret.
CELL_ROWS_PER_BLOCK = 17
CELL_COLS_PER_BLOCK = 18

# Bloklar arası "koridor" boşluğu (grid kolon/satır olarak)
GAP_R = 2
GAP_C = 2


def main() -> None:
    total = END_LOOM - START_LOOM + 1

    cells = []
    loom = START_LOOM

    for br in range(BLOCK_ROWS):
        for bc in range(BLOCK_COLS):
            origin_r = br * (CELL_ROWS_PER_BLOCK + GAP_R)
            origin_c = bc * (CELL_COLS_PER_BLOCK + GAP_C)

            for r in range(CELL_ROWS_PER_BLOCK):
                for c in range(CELL_COLS_PER_BLOCK):
                    if loom > END_LOOM:
                        break
                    cells.append({"loom": loom, "row": origin_r + r, "col": origin_c + c})
                    loom += 1

            if loom > END_LOOM:
                break
        if loom > END_LOOM:
            break

    produced = len(cells)
    if produced != total:
        raise RuntimeError(
            f"Produced {produced} cells but expected {total}. "
            f"Adjust CELL_ROWS_PER_BLOCK/CELL_COLS_PER_BLOCK."
        )

    data = {
        "site": "ISKO11",
        "schema": "cells-v1",
        "global_range": [START_LOOM, END_LOOM],
        "cells": cells,
        # divider'ları şimdilik boş bırakıyoruz; istersek sonra ekleriz
        "dividers_v": [],
        "dividers_h": [],
    }

    OUT.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Wrote: {OUT}  (cells={produced})")


if __name__ == "__main__":
    main()
