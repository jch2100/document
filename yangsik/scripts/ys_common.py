"""양식대로 공통 유틸: 테마/프로필 로딩, 경로, 텍스트 마크업."""
from __future__ import annotations

import json
import re
import sys
from pathlib import Path

try:
    import yaml
except ImportError:  # pragma: no cover
    sys.exit("PyYAML이 필요합니다: pip install pyyaml")

SKILL_DIR = Path(__file__).resolve().parent.parent
TEMPLATES_DIR = SKILL_DIR / "templates"
RULES_DIR = SKILL_DIR / "rules"
THEME_DEFAULT = SKILL_DIR / "theme" / "default.yaml"

EMU_PER_INCH = 914400


def inch(v: float) -> int:
    return int(round(v * EMU_PER_INCH))


def load_yaml(path) -> dict:
    with open(path, encoding="utf-8") as f:
        return yaml.safe_load(f) or {}


def dump_yaml(data, path) -> None:
    with open(path, "w", encoding="utf-8") as f:
        yaml.safe_dump(data, f, allow_unicode=True, sort_keys=False, width=100)


def load_json(path) -> dict:
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def load_theme(path=None) -> dict:
    theme = load_yaml(THEME_DEFAULT)
    if path:
        user = load_yaml(path)
        for k, v in user.items():
            if isinstance(v, dict) and isinstance(theme.get(k), dict):
                theme[k].update(v)
            else:
                theme[k] = v
    for k, v in theme["colors"].items():
        theme["colors"][k] = str(v).lstrip("#").upper()
    return theme


BOLD_RE = re.compile(r"(\*\*.+?\*\*)")


def split_bold(text: str):
    """'앞 **강조** 뒤' -> [('앞 ', False), ('강조', True), (' 뒤', False)]"""
    out = []
    for part in BOLD_RE.split(str(text)):
        if not part:
            continue
        if part.startswith("**") and part.endswith("**") and len(part) > 4:
            out.append((part[2:-2], True))
        else:
            out.append((part, False))
    return out


def flatten_items(items, level=0):
    """['a', ['a-1', 'a-2'], 'b'] -> [(0,'a'), (1,'a-1'), (1,'a-2'), (0,'b')]"""
    if isinstance(items, str):
        return [(level, items)]
    out = []
    for it in items or []:
        if isinstance(it, (list, tuple)):
            out.extend(flatten_items(it, level + 1))
        elif isinstance(it, dict):  # {"text": "...", "sub": [...]}
            out.append((level, it.get("text", "")))
            out.extend(flatten_items(it.get("sub", []), level + 1))
        else:
            out.append((level, str(it)))
    return out


def find_template(kind: str, name: str) -> Path:
    """kind: pptx|docx, name: '01-exec-report' 또는 '01' 또는 경로."""
    p = Path(name)
    if p.exists():
        return p
    base = TEMPLATES_DIR / kind
    for d in sorted(base.iterdir()):
        if d.is_dir() and (d.name == name or d.name.startswith(name + "-") or d.name.split("-", 1)[-1] == name):
            return d / f"template.{kind}"
    sys.exit(f"양식을 찾을 수 없습니다: {kind}/{name}. 목록: python scripts/list_templates.py")
