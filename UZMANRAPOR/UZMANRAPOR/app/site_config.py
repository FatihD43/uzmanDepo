# app/site_config.py
from __future__ import annotations

import os
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

# ---------------------------------------------------------------------
# Site seçimi
# ---------------------------------------------------------------------
# Env:
#   UZMANRAPOR_SITE = ISKO14 | ISKO11 | MEKIKLI
# Yoksa varsayılan ISKO14.
DEFAULT_SITE = "ISKO14"


def _env(name: str, default: str = "") -> str:
    v = os.getenv(name)
    return v.strip() if v else default


def get_site() -> str:
    site = _env("UZMANRAPOR_SITE", DEFAULT_SITE).upper()
    if site not in ("ISKO14", "ISKO11", "MEKIKLI"):
        # Hatalı değer gelirse sistemi kırmayalım; ISKO14'e düş.
        return DEFAULT_SITE
    return site


# ---------------------------------------------------------------------
# Site konfigürasyon modeli
# ---------------------------------------------------------------------
@dataclass(frozen=True)
class SiteConfig:
    key: str
    db_name: str
    loom_start: int
    loom_end: int
    # UI'de gösterilecek kategori listesi (örn: ["Tümü","DENIM","HAM"])
    categories: List[str]
    # HAM aralıkları (dahil). DENIM aralığı yazmayacağız; DENIM = range içinde HAM olmayanlar.
    ham_ranges: List[Tuple[int, int]]


SITES: Dict[str, SiteConfig] = {
    "ISKO14": SiteConfig(
        key="ISKO14",
        db_name="UzmanRaporDB_ISKO14",
        loom_start=2201,
        loom_end=2518,
        categories=["Tümü", "DENIM", "HAM"],
        # ISKO14 HAM aralığı (mevcut sistemde 2447–2518 olarak kullanıyorduk)
        ham_ranges=[(2447, 2518)],
    ),
    "ISKO11": SiteConfig(
    key="ISKO11",
    db_name="UzmanRaporDB_ISKO11",
    loom_start=1301,
    loom_end=1912,
    categories=["Tümü"],   # ISKO11’de DENIM/HAM yok
    ham_ranges=[],
    ),

    "MEKIKLI": SiteConfig(
        key="MEKIKLI",
        db_name="UzmanRaporDB_MEKIKLI",
        loom_start=601,
        loom_end=700,
        categories=["Tümü", "DENIM", "HAM"],
        # Mekikli HAM aralığı: 633–648 (dahil)
        ham_ranges=[(633, 648)],
    ),
}


def get_site_config(site: Optional[str] = None) -> SiteConfig:
    key = (site or get_site()).upper()
    return SITES.get(key, SITES[DEFAULT_SITE])


# ---------------------------------------------------------------------
# Ortak yardımcılar
# ---------------------------------------------------------------------
def get_db_name(site: Optional[str] = None) -> str:
    return get_site_config(site).db_name


def get_loom_range(site: Optional[str] = None) -> Tuple[int, int]:
    cfg = get_site_config(site)
    return cfg.loom_start, cfg.loom_end


def get_categories(site: Optional[str] = None) -> List[str]:
    return list(get_site_config(site).categories)


def _in_any_ranges(value: int, ranges: List[Tuple[int, int]]) -> bool:
    for a, b in ranges:
        if a <= value <= b:
            return True
    return False


def loom_in_category(loom_no: int, category: str, site: Optional[str] = None) -> bool:
    """
    category:
      - "Tümü"  -> site loom aralığındaki her tezgah
      - "HAM"   -> ham_ranges içinde
      - "DENIM" -> site aralığında olup HAM olmayanlar
    ISKO11 gibi sites'ta categories=["Tümü"] ise DENIM/HAM sorgulansa bile güvenli şekilde davranır.
    """
    cfg = get_site_config(site)
    cat = (category or "Tümü").upper()

    # Önce site aralığı kontrolü
    if not (cfg.loom_start <= loom_no <= cfg.loom_end):
        return False

    if cat in ("TÜMÜ", "TUMU", "ALL"):
        return True

    # Eğer bu site DENIM/HAM kullanmıyorsa:
    if cfg.categories == ["Tümü"]:
        return True

    is_ham = _in_any_ranges(loom_no, cfg.ham_ranges)

    if cat == "HAM":
        return is_ham
    if cat == "DENIM":
        return not is_ham

    # Bilinmeyen kategori -> kırma, "Tümü" gibi davran
    return True
