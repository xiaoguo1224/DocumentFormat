#!/usr/bin/env python3
from __future__ import annotations

from pathlib import Path
from typing import Any
import yaml

DEFAULT_PROFILE = Path(__file__).resolve().parents[1] / 'templates/generic-formal/profile.yaml'


def load_profile(path: Path | None = None) -> dict[str, Any]:
    target = path or DEFAULT_PROFILE
    with target.open('r', encoding='utf-8') as f:
        data = yaml.safe_load(f) or {}
    data.setdefault('_profile_path', str(target))
    return data


def get(profile: dict[str, Any], *keys: str, default=None):
    cur: Any = profile
    for key in keys:
        if not isinstance(cur, dict) or key not in cur:
            return default
        cur = cur[key]
    return cur


def style(profile: dict[str, Any], role: str, fallback: str) -> str:
    return str(get(profile, 'styles', role, default=fallback) or fallback)


def normalized_titles(profile: dict[str, Any], section: str, defaults: list[str]) -> set[str]:
    values = get(profile, 'sections', 'titles', section, default=None)
    if values is None:
        values = defaults
    if isinstance(values, str):
        values = [values]
    return {''.join(str(v).split()).casefold() for v in values}
