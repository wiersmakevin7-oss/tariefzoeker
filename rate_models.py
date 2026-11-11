"""Stable, importable types for Streamlit's pickle-backed tariff cache."""
from dataclasses import dataclass


@dataclass(frozen=True)
class RateRow:
    country: str
    zone: str
    max_kg: float | None
    max_ldm: float | None
    max_eu_pallets: float | None
    max_block_pallets: float | None
    base_price: float
    transit: str | None
    carrier: str
    source: str
