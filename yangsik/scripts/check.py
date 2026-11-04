#!/usr/bin/env python3
"""검증 게이트: PPTX/DOCX 를 한국어 문서 기준으로 검사하고 리포트를 만듭니다.

사용:
  python scripts/check.py 결과.pptx                 # 화면 요약 + 결과_check.md
  python scripts/check.py 결과.docx --profile profiles/우리회사/profile.yaml
  python scripts/check.py 결과.pptx --json          # 기계가 읽는 결과

판정: FAIL 이 하나라도 있으면 종료코드 1.
검사 항목
  양식    off_template_shape(양식 밖 텍스트박스) · empty_slot · guide_left([안내문] 남음)
  글꼴    missing_korean_font · korean_lang · line_break(어절 단위 줄바꿈 설정)
  넘침    overflow_high / overflow_maybe (글자 폭 추정치) · table_tall
  문체    ai_tone · ending_style(개조식 종결) · item_long · gongmun(날짜 등 표기)
  Word    direct_symbol(기호 직접 입력) · direct_format · empty_paragraphs · heading_number_dup · compat_toc
"""
from __future__ import annotations

import argparse
import json
import math
import re
import sys
import zipfile
from dataclasses import dataclass, asdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from ys_common import RULES_DIR, load_yaml  # noqa: E402

HANGUL = re.compile(r"[가-힣ㄱ-ㆎ]")
GUIDE = re.compile(r"\[[^\]\n]{1,60}\]")
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

LEVEL_ORDER = {"FAIL": 0, "WARN": 1, "INFO": 2}
DEFAULT_LEVEL = {
    "off_template_shape": "FAIL", "off_template_object": "WARN", "missing_korean_font": "FAIL",
    "korean_lang": "INFO", "line_break": "WARN", "overflow_high": "FAIL", "overflow_maybe": "WARN",
    "table_tall": "WARN", "empty_slot": "WARN", "guide_left": "WARN", "ai_tone": "WARN",
    "ending_style": "WARN", "item_long": "INFO", "gongmun": "INFO", "direct_symbol": "WARN",
    "direct_format": "WARN", "empty_paragraphs": "WARN", "heading_number_dup": "WARN",
    "table_no_style": "INFO", "compat_toc": "INFO", "topic_title": "WARN", "title_long": "WARN", "missing_source": "WARN",
}
LABEL = {
    "off_template_shape": "양식 밖 텍스트", "off_template_object": "양식 밖 개체", "missing_korean_font": "한글 글꼴",
    "korean_lang": "언어 태그", "line_break": "줄바꿈 설정", "overflow_high": "넘침(확실)", "overflow_maybe": "넘침(가능)",
    "table_tall": "표 높이", "empty_slot": "빈 슬롯", "guide_left": "안내문 남음", "ai_tone": "AI 문체",
    "ending_style": "개조식 종결", "item_long": "긴 항목", "gongmun": "표기", "direct_symbol": "기호 직접 입력",
    "direct_format": "직접 서식", "empty_paragraphs": "빈 줄 반복", "heading_number_dup": "번호 중복",
    "table_no_style": "표 스타일", "compat_toc": "호환성", "topic_title": "주제어 제목", "title_long": "긴 제목",
    "missing_source": "출처 없음",
}


@dataclass
class Finding:
    code: str
    where: str
    msg: str
    fix: str = ""
    level: str = ""


class Checker:
    def __init__(self, profile=None):
        self.findings: list[Finding] = []
        self.ai = load_yaml(RULES_DIR / "ai-tone.yaml")
        self.gj = load_yaml(RULES_DIR / "gaejosik.yaml")
        self.gm = load_yaml(RULES_DIR / "gongmun.yaml")
        self.ty = load_yaml(RULES_DIR / "typography.yaml")
        self.levels = dict(DEFAULT_LEVEL)
        self.profile = profile or {}
        chk = self.profile.get("check", {})
        for code in chk.get("fail_on", []):
            self.levels[code] = "FAIL"
        for code in chk.get("warn_on", []):
            if self.levels.get(code) != "FAIL":
                self.levels[code] = "WARN"
        for code in chk.get("ignore", []):
            self.levels[code] = "OFF"
        banned = self.profile.get("writing", {}).get("banned", [])
        self.extra_banned = [b for b in banned if b not in ("—", "이모지") and not b.startswith("~")]
        self.extra_banned += [b[1:] for b in banned if b.startswith("~")]

    def add(self, code, where, msg, fix="", level=None):
        lv = level.upper() if level else self.levels.get(code, "WARN")
        if lv == "OFF":
            return
        self.findings.append(Finding(code, where, msg, fix, lv))

    # ---------------- 문체 ----------------
    def text_rules(self, text, where, bulletish=False, allow_endings=False):
        t = text.strip()
        if not t:
            return
        for pat in self.ai["patterns"]:
            if re.search(pat["re"], t, flags=re.M):
                self.add("ai_tone", where, f'{pat["msg"]}: "{_clip(t)}"', pat.get("fix", ""), pat.get("level"))
        for word in self.extra_banned:
            if word and word in t:
                self.add("ai_tone", where, f'금지어 "{word}": "{_clip(t)}"', "profile.yaml writing.banned")
        for pat in self.gm["patterns"]:
            m = re.search(pat["re"], t)
            if m and not (pat.get("exclude_re") and re.search(pat["exclude_re"], t)):
                self.add("gongmun", where, f'{pat["msg"]}: "{m.group(0)}"', pat.get("fix", ""), pat.get("level"))
        if GUIDE.search(t):
            self.add("guide_left", where, f'양식 안내문이 남아 있음: "{GUIDE.search(t).group(0)}"', "실제 내용으로 바꾸거나 삭제")
        if bulletish and not allow_endings:
            if not any(p in t for p in self.gj.get("allow_phrases", [])):
                for pat in self.gj["bad_endings"]:
                    if re.search(pat["re"], t):
                        self.add("ending_style", where, f'{pat["msg"]}: "{_clip(t)}"', pat.get("fix", ""), pat.get("level"))
                        break
            if len(t) > self.gj.get("max_item_chars", 80):
                self.add("item_long", where, f"항목이 {len(t)}자 (권장 {self.gj['max_item_chars']}자 이내)", "두 항목으로 나누거나 하위 항목으로")

    def is_korean_font(self, name):
        if not name:
            return False
        return any(k.lower() in name.lower() for k in self.ty["korean_fonts"]) or bool(HANGUL.search(name))

    # ---------------- 넘침 추정 ----------------
    def text_width_em(self, s):
        cw = self.ty["char_width"]
        w = 0.0
        for ch in s:
            if HANGUL.search(ch):
                w += cw["hangul"]
            elif "　" <= ch <= "鿿" or "＀" <= ch <= "￯":
                w += cw["cjk"]
            elif ch.isdigit():
                w += cw["digit"]
            elif ch.isupper():
                w += cw["latin_upper"]
            elif ch.islower():
                w += cw["latin_lower"]
            elif ch == " ":
                w += cw["space"]
            elif ch in ".,:;'\"()[]-·/!?":
                w += cw["punct"]
            else:
                w += cw["other"]
        return w

    def report(self):
        nat = lambda w: [int(x) if x.isdigit() else x for x in re.split(r"(\d+)", w)]  # noqa: E731
        self.findings.sort(key=lambda f: (LEVEL_ORDER[f.level], nat(f.where)))
        n = {k: sum(1 for f in self.findings if f.level == k) for k in LEVEL_ORDER}
        return n


class _LazyCK:
    _ck = None

    def __getattr__(self, k):
        if _LazyCK._ck is None:
            _LazyCK._ck = Checker()
        return getattr(_LazyCK._ck, k)


_CK = _LazyCK()


def _clip(t, n=40):
    t = t.replace("\n", " ")
    return t if len(t) <= n else t[: n - 1] + "…"


# =====================================================================
# PPTX
# =====================================================================
def _lvl_ppr(lststyle, level):
    if lststyle is None:
        return None
    return lststyle.find(f"{A}lvl{level + 1}pPr")


def _inherit_chain(shape, prs):
    """슬라이드 도형 → 레이아웃 플레이스홀더 → 마스터 플레이스홀더 의 txBody 목록 + 마스터 스타일."""
    chain = []
    el = shape._element
    chain.append(el)
    if shape.is_placeholder:
        try:
            base = shape._base_placeholder
        except Exception:
            base = None
        if base is not None:
            chain.append(base._element)
            try:
                mb = base._base_placeholder
                if mb is not None:
                    chain.append(mb._element)
            except Exception:
                pass
    return chain


def _resolve(chain, prs, level, attr, is_title):
    """attr: 'sz' (defRPr) | 'lnSpc' | 'spcBef' | 'marL'."""
    for el in chain:
        lst = el.find(f".//{A}lstStyle")
        ppr = _lvl_ppr(lst, level)
        v = _from_ppr(ppr, attr)
        if v is not None:
            return v
    master = prs.slide_masters[0]._element
    sty = master.find(f"{P}txStyles/{P}{'titleStyle' if is_title else 'bodyStyle'}")
    v = _from_ppr(_lvl_ppr(sty, level), attr)
    if v is not None:
        return v
    return {"sz": 18.0, "lnSpc": None, "spcBef": 0.0, "marL": 0.0}[attr]


def _from_ppr(ppr, attr):
    if ppr is None:
        return None
    if attr == "sz":
        d = ppr.find(f"{A}defRPr")
        if d is not None and d.get("sz"):
            return int(d.get("sz")) / 100
    elif attr == "lnSpc":
        x = ppr.find(f"{A}lnSpc/{A}spcPct")
        if x is not None:
            return int(x.get("val")) / 100000
    elif attr == "spcBef":
        x = ppr.find(f"{A}spcBef/{A}spcPts")
        if x is not None:
            return int(x.get("val")) / 100
    elif attr == "marL":
        if ppr.get("marL") is not None:
            return int(ppr.get("marL")) / 914400
    return None


def _insets(chain):
    d = {"lIns": 0.1, "rIns": 0.1, "tIns": 0.05, "bIns": 0.05}
    found = {}
    for el in reversed(chain):
        bp = el.find(f".//{A}bodyPr")
        if bp is not None:
            for k in d:
                if bp.get(k) is not None:
                    found[k] = int(bp.get(k)) / 914400
    d.update(found)
    return d


def _autofit(chain):
    for el in chain:
        bp = el.find(f".//{A}bodyPr")
        if bp is not None:
            if bp.find(f"{A}normAutofit") is not None:
                fs = bp.find(f"{A}normAutofit").get("fontScale")
                return "norm", (int(fs) / 100000 if fs else 1.0)
            if bp.find(f"{A}spAutoFit") is not None:
                return "sp", 1.0
            if bp.find(f"{A}noAutofit") is not None:
                return "none", 1.0
    return "none", 1.0


def check_pptx(path, ck: Checker):
    from pptx import Presentation
    from pptx.util import Emu
    prs = Presentation(str(path))
    slide_h = prs.slide_height / 914400

    # 글꼴: 테마
    master = prs.slide_masters[0]
    from pptx.opc.constants import RELATIONSHIP_TYPE as RT
    from lxml import etree
    theme = etree.fromstring(master.part.part_related_by(RT.THEME).blob)
    for kind in ("majorFont", "minorFont"):
        f = theme.find(f".//{A}{kind}")
        ea = f.find(f"{A}ea").get("typeface") if f is not None and f.find(f"{A}ea") is not None else ""
        hang = next((x.get("typeface") for x in f.findall(f"{A}font") if x.get("script") == "Hang"), "") if f is not None else ""
        latin = f.find(f"{A}latin").get("typeface") if f is not None else ""
        if not (ck.is_korean_font(ea) or ck.is_korean_font(hang) or ck.is_korean_font(latin)):
            ck.add("missing_korean_font", f"테마 {kind}", f"테마 글꼴에 한글 글꼴이 지정되지 않음 (ea='{ea}', Hang='{hang}')",
                   "normalize.py --korean-font '맑은 고딕' 으로 테마의 동아시아 글꼴 지정")
    body = master._element.find(f"{P}txStyles/{P}bodyStyle/{A}lvl1pPr")
    if body is not None and body.get("latinLnBrk") != "0":
        ck.add("line_break", "마스터", "어절 단위 줄바꿈이 꺼져 있음 → 한글 단어가 중간에서 잘릴 수 있음",
               "normalize.py 실행 (PowerPoint: 단락 → 한글 입력 체계 → '한글 단어 잘림 허용' 해제)")

    NON_ACTION = {"표지", "섹션", "목차", "마무리", "핵심메시지"}
    for s_i, slide in enumerate(prs.slides, 1):
        layout_name = slide.slide_layout.name
        allow_end = layout_name in ck.gj.get("allow_layouts", [])
        # 액션 타이틀 · 출처 (컨설팅 원칙)
        title = next((s for s in slide.placeholders if "TITLE" in str(s.placeholder_format.type)), None)
        if title is not None and layout_name not in NON_ACTION and title.text_frame.text.strip():
            tt = title.text_frame.text.strip()
            if len(tt) < 12 and not re.search(r"\d", tt):
                ck.add("topic_title", f"슬라이드 {s_i} · 제목", f'제목이 주제어("{tt}") — 이 장의 결론 문장으로',
                       "예: '현황' → '이용률은 3년째 하락, 현장직 접속이 막혀 있음'")
            elif len(tt) > 55:
                ck.add("title_long", f"슬라이드 {s_i} · 제목", f"제목 {len(tt)}자 — 2줄을 넘길 수 있음", "45자 안팎으로 줄이고 세부는 본문으로")
        has_data = any((getattr(s, "has_chart", False) and s.has_chart) or (getattr(s, "has_table", False) and s.has_table
                       and re.search(r"\d", " ".join(c.text for r in s.table.rows for c in r.cells))) for s in slide.shapes)
        src = next((s for s in slide.placeholders if s.name == "출처"), None)
        if has_data and (src is None or not src.text_frame.text.replace("출처:", "").strip()):
            ck.add("missing_source", f"슬라이드 {s_i}", "숫자가 있는 표·차트에 출처가 없음", "'출처' 슬롯에 자료·기준일 (예: 출처: 인사시스템, 2026. 10.)")
        for shape in _iter_shapes(slide.shapes):
            where = f"슬라이드 {s_i} · {shape.name}"
            if not shape.is_placeholder:
                if shape.has_text_frame and shape.text_frame.text.strip():
                    ck.add("off_template_shape", where, f'레이아웃 슬롯이 아닌 텍스트박스: "{_clip(shape.text_frame.text)}"',
                           "레이아웃의 플레이스홀더에 넣기 (양식 변경 시 따라가지 않음)")
                elif (getattr(shape, "has_chart", False) and shape.has_chart) or (getattr(shape, "has_table", False) and shape.has_table):
                    ck.add("off_template_object", where, "레이아웃 슬롯 밖의 표/차트", "표·차트 레이아웃 사용")
            if getattr(shape, "has_table", False) and shape.has_table:
                _check_table(shape, slide_h, where, ck)
                continue
            if not shape.has_text_frame:
                continue
            tf = shape.text_frame
            txt = tf.text.strip()
            if shape.is_placeholder and not txt:
                ck.add("empty_slot", where, "비어 있는 슬롯 (편집 화면에 안내문이 보임)", "내용을 넣거나 슬롯 삭제")
                continue
            is_title = shape.is_placeholder and "TITLE" in str(shape.placeholder_format.type)
            # 문체
            bulletish = shape.is_placeholder and not is_title and ("본문" in shape.name or "카드" in shape.name or "Content" in shape.name or "Text" in shape.name)
            for p in tf.paragraphs:
                t = "".join(r.text for r in p.runs)
                ck.text_rules(t, where, bulletish=bulletish, allow_endings=allow_end)
                for r in p.runs:
                    if HANGUL.search(r.text):
                        rpr = r._r.find(f"{A}rPr")
                        lang = rpr.get("lang") if rpr is not None else None
                        if lang and not lang.startswith("ko"):
                            ck.add("korean_lang", where, f"한글 텍스트의 언어가 {lang}", "lang='ko-KR' (맞춤법·줄바꿈 규칙)")
                            break
                        latin = rpr.find(f"{A}latin") if rpr is not None else None
                        ea = rpr.find(f"{A}ea") if rpr is not None else None
                        if latin is not None and ea is None and not ck.is_korean_font(latin.get("typeface")) and not latin.get("typeface", "").startswith("+"):
                            ck.add("missing_korean_font", where, f"영문 글꼴 '{latin.get('typeface')}'만 지정, 한글 글꼴 없음",
                                   "동아시아 글꼴(ea) 지정 또는 직접 서식 제거")
                            break
            # 넘침
            _check_overflow(shape, prs, is_title, where, ck)


def _iter_shapes(shapes):
    for s in shapes:
        if "GROUP" in str(s.shape_type):
            yield from _iter_shapes(s.shapes)
        else:
            yield s


def _text_metrics(shape, prs, is_title, ty, ck=None):
    """(필요 높이 in, 상자 높이 in, 자동맞춤 모드) — 글자 폭 추정."""
    ck = ck or _CK
    chain = _inherit_chain(shape, prs)
    ins = _insets(chain)
    box_w = shape.width / 914400 - ins["lIns"] - ins["rIns"]
    box_h = shape.height / 914400 - ins["tIns"] - ins["bIns"]
    mode, scale = _autofit(chain)
    total = 0.0
    for i, p in enumerate(shape.text_frame.paragraphs):
        lvl = p.level
        sz = next((r.font.size.pt for r in p.runs if r.font.size), None)
        sz = (sz or _resolve(chain, prs, lvl, "sz", is_title)) * scale
        ln = (_resolve(chain, prs, lvl, "lnSpc", is_title) or 1.0) * ty["line_height"]
        bef = _resolve(chain, prs, lvl, "spcBef", is_title) if i else 0.0
        marL = _resolve(chain, prs, lvl, "marL", is_title) or 0.0
        avail = max(box_w - marL, 0.5)
        text = "".join(r.text for r in p.runs)
        em_w = ck.text_width_em(text) * sz / 72
        # 굵은 글자는 약 5% 넓고, 줄 끝 어절이 통째로 넘어가므로 폭의 94%만 쓴다고 본다
        bold = any(r.font.bold for r in p.runs) or _resolve_bold(chain, lvl)
        em_w *= 1.05 if bold else 1.0
        lines = max(1, math.ceil(em_w / (avail * 0.94))) if text else 1
        total += lines * sz * ln / 72 + bef / 72
    return total, box_h, mode, ins


def _resolve_bold(chain, level):
    for el in chain:
        ppr = _lvl_ppr(el.find(f".//{A}lstStyle"), level)
        if ppr is not None:
            d = ppr.find(f"{A}defRPr")
            if d is not None and d.get("b") is not None:
                return d.get("b") == "1"
    return False


def estimate_height(shape, prs, ty=None):
    """채우기 후 배치용: 글이 차지할 높이 + 상하 안쪽 여백 (인치)."""
    ty = ty or _CK.ty
    is_title = shape.is_placeholder and "TITLE" in str(shape.placeholder_format.type)
    need, _box, _mode, ins = _text_metrics(shape, prs, is_title, ty)
    return need + ins["tIns"] + ins["bIns"] + 0.05


def _check_overflow(shape, prs, is_title, where, ck: Checker):
    if shape.width is None or shape.height is None:
        return
    total, box_h, mode, _ins = _text_metrics(shape, prs, is_title, ck.ty, ck)
    if box_h <= 0:
        return
    ratio = total / box_h
    if ratio > ck.ty["overflow_fail"]:
        if mode == "norm":
            ck.add("overflow_maybe", where, f"글이 상자보다 약 {ratio:.1f}배 김 → 자동 맞춤으로 글자가 작게 줄어듦",
                   "항목 줄이기 또는 장 나누기")
        else:
            ck.add("overflow_high", where, f"글이 상자보다 약 {ratio:.1f}배 김 (넘침 거의 확실)", "항목 줄이기 또는 장 나누기")
    elif ratio > ck.ty["overflow_warn"]:
        ck.add("overflow_maybe", where, f"상자를 꽉 채움 (추정 {ratio:.2f}배)", "한두 줄 줄이기 권장")


def _check_table(gf, slide_h, where, ck: Checker):
    table = gf.table
    top = gf.top / 914400
    est = 0.0
    for r_i, row in enumerate(table.rows):
        row_h = row.height / 914400
        need = 0.0
        for c_i, cell in enumerate(row.cells):
            w = table.columns[c_i].width / 914400 - 0.2
            lines = 0
            for p in cell.text_frame.paragraphs:
                t = "".join(r.text for r in p.runs)
                sz = next((r.font.size.pt for r in p.runs if r.font.size), 12)
                lines += max(1, math.ceil(ck.text_width_em(t) * sz / 72 / max(w, 0.3)))
                ck.text_rules(t, f"{where} · 표({r_i+1},{c_i+1})")
            need = max(need, lines * sz * 1.2 / 72 + 0.1)
        est += max(row_h, need)
    if top + est > slide_h - 0.55:
        ck.add("table_tall", where, f"표 예상 높이 {est:.1f}in → 하단 여백/꼬리말 침범 가능", "행 줄이기, 글 줄이기, 표 나누기")


# =====================================================================
# DOCX
# =====================================================================
def check_docx(path, ck: Checker):
    from docx import Document
    from lxml import etree
    doc = Document(str(path))
    # 글꼴
    styles_el = doc.styles.element
    rf = styles_el.find(f"{W}docDefaults/{W}rPrDefault/{W}rPr/{W}rFonts")
    ea = rf.get(f"{W}eastAsia") if rf is not None else None
    ea_theme = rf.get(f"{W}eastAsiaTheme") if rf is not None else None
    ok = ck.is_korean_font(ea)
    if not ok and ea_theme:
        with zipfile.ZipFile(path) as z:
            names = [n for n in z.namelist() if n.startswith("word/theme/")]
            if names:
                th = etree.fromstring(z.read(names[0]))
                kind = "majorFont" if "major" in ea_theme else "minorFont"
                f = th.find(f".//{A}{kind}")
                cands = [f.find(f"{A}ea").get("typeface") if f.find(f"{A}ea") is not None else ""]
                cands += [x.get("typeface") for x in f.findall(f"{A}font") if x.get("script") == "Hang"]
                ok = any(ck.is_korean_font(c) for c in cands)
    if not ok:
        nrf = doc.styles["Normal"].element.find(f"{W}rPr/{W}rFonts")
        ok = nrf is not None and ck.is_korean_font(nrf.get(f"{W}eastAsia"))
    if not ok:
        ck.add("missing_korean_font", "스타일 기본값", "문서 기본 한글(동아시아) 글꼴이 지정되지 않음",
               "normalize.py --korean-font '맑은 고딕'")

    bullet_styles = {"개조식 1", "개조식 2", "개조식 3", "List Bullet", "List Bullet 2", "List Bullet 3", "List Paragraph"}
    prof_styles = ck.profile.get("styles", {})
    bullet_styles |= {v for k, v in prof_styles.items() if k.startswith("bullet")}
    heading_styles = {"Heading 1", "Heading 2", "Heading 3"} | {v for k, v in prof_styles.items() if k in ("h1", "h2", "h3")}
    manual = re.compile(ck.gj["manual_symbols"])
    head_num = re.compile(r"^\s*(?:\d+\.|[가나다라마바사아자차카타파하]\.|\d+\)|[ⅠⅡⅢⅣⅤ]+\.)\s*")

    runs_total = runs_direct = 0
    empty_streak = 0
    paragraphs = list(doc.paragraphs)
    for i, p in enumerate(paragraphs, 1):
        where = f"문단 {i}"
        style = p.style.name if p.style is not None else "Normal"
        t = p.text
        has_num = p._p.pPr is not None and p._p.pPr.numPr is not None
        has_break = p._p.find(f".//{W}br") is not None or p._p.find(f".//{W}sectPr") is not None
        if not t.strip():
            empty_streak = empty_streak + 1 if not has_break else 0
            if empty_streak == 2:
                ck.add("empty_paragraphs", where, "빈 문단을 연달아 넣어 간격을 띄움", "문단 '앞/뒤 간격' 또는 쪽 나눔 사용")
            continue
        empty_streak = 0
        where = f"문단 {i} ({style})"
        is_bullet = style in bullet_styles or (has_num and style not in heading_styles) or bool(re.match(r"^\s*[□■○●◦\-–·•▪]\s", t))
        if manual.match(t) and style not in heading_styles | SYMBOL_OK_STYLES and not has_num:
            ck.add("direct_symbol", where, f'기호/번호를 직접 입력: "{_clip(t)}"', "'개조식 1~3' 스타일 적용 후 기호 삭제 (normalize.py --fix-symbols)")
        if style in heading_styles and head_num.match(t) and _style_has_num(doc, style):
            ck.add("heading_number_dup", where, f'제목 자동 번호 + 직접 입력 번호: "{_clip(t)}"', "텍스트의 번호 삭제")
        ck.text_rules(t, where, bulletish=is_bullet)
        for r in p.runs:
            runs_total += 1
            rpr = r._r.rPr
            if rpr is not None and (rpr.find(f"{W}sz") is not None or rpr.find(f"{W}rFonts") is not None):
                runs_direct += 1
    for t_i, table in enumerate(doc.tables, 1):
        tblStyle = table._tbl.tblPr.find(f"{W}tblStyle")
        if tblStyle is None:
            ck.add("table_no_style", f"표 {t_i}", "표 스타일 없음 (직접 서식 표)", "'양식대로 표' 스타일 적용")
        for r_i, row in enumerate(table.rows, 1):
            seen = set()
            for c_i, cell in enumerate(row.cells, 1):
                if id(cell._tc) in seen:
                    continue
                seen.add(id(cell._tc))
                ck.text_rules(cell.text, f"표 {t_i}({r_i},{c_i})")
    if runs_total >= 20 and runs_direct / runs_total > 0.3:
        ck.add("direct_format", "문서 전체", f"글자 {runs_direct}/{runs_total}개 조각에 글꼴·크기를 직접 지정 ({runs_direct/runs_total:.0%})",
               "스타일로 서식 통일 (Word: Ctrl+Space 로 직접 서식 해제)")
    if doc.element.body.find(f".//{W}instrText") is not None and any("TOC" in (x.text or "") for x in doc.element.body.iter(f"{W}instrText")):
        ck.add("compat_toc", "목차", "목차 필드: Word에서 '필드 업데이트' 필요, Google 문서에서는 비어 보일 수 있음", "Word에서 열어 F9")


SYMBOL_OK_STYLES = {"Caption", "캡션", "표 본문", "Title", "Subtitle", "요약 박스", "참고", "표지 정보"}


def _style_has_num(doc, style_name):
    try:
        st = doc.styles[style_name]
    except KeyError:
        return False
    ppr = st.element.pPr
    return ppr is not None and ppr.numPr is not None


# =====================================================================
def write_report(path, ck: Checker, out=None):
    n = ck.report()
    verdict = "FAIL" if n["FAIL"] else ("WARN" if n["WARN"] else "PASS")
    icon = {"FAIL": "❌", "WARN": "⚠️", "INFO": "ℹ️", "PASS": "✅"}
    lines = [f"# 검증 리포트: {Path(path).name}", "",
             f"**판정: {icon[verdict]} {verdict}**  ·  실패 {n['FAIL']} · 경고 {n['WARN']} · 참고 {n['INFO']}", ""]
    if verdict == "PASS":
        lines.append("양식·글꼴·넘침·문체 검사를 모두 통과했습니다. (넘침은 글자 폭 추정치입니다)")
    by_code = {}
    for f in ck.findings:
        by_code.setdefault((f.level, f.code), []).append(f)
    if ck.findings:
        lines += ["## 요약", "", "| 수준 | 항목 | 건수 | 고치는 법 |", "|---|---|---|---|"]
        for (lv, code), fs in sorted(by_code.items(), key=lambda kv: (LEVEL_ORDER[kv[0][0]], kv[0][1])):
            lines.append(f"| {icon[lv]} {lv} | {LABEL.get(code, code)} | {len(fs)} | {fs[0].fix} |")
        lines += ["", "## 상세", "", "| 수준 | 위치 | 내용 |", "|---|---|---|"]
        for f in ck.findings:
            lines.append(f"| {f.level} | {f.where} | {f.msg.replace('|', '/')} |")
    lines += ["", "---", "넘침 판정은 글자 폭 추정치입니다. 최종 확인은 PowerPoint/Word에서 하세요."]
    out = Path(out) if out else Path(path).with_name(Path(path).stem + "_" + Path(path).suffix.lstrip(".") + "_check.md")
    out.write_text("\n".join(lines), encoding="utf-8")
    return verdict, n, out


def run(path, profile=None, out=None, quiet=False):
    prof = load_yaml(profile) if profile else None
    ck = Checker(prof)
    ext = Path(path).suffix.lower()
    if ext == ".pptx":
        check_pptx(path, ck)
    elif ext == ".docx":
        check_docx(path, ck)
    else:
        sys.exit("pptx 또는 docx 파일만 검사합니다")
    verdict, n, rpt = write_report(path, ck, out)
    if not quiet:
        print(f"[검증] {Path(path).name}: {verdict}  (실패 {n['FAIL']} · 경고 {n['WARN']} · 참고 {n['INFO']})  → {rpt}")
        for f in [f for f in ck.findings if f.level != "INFO"][:12]:
            print(f"  {f.level:4} {f.where}: {f.msg}")
        more = len([f for f in ck.findings if f.level != "INFO"]) - 12
        if more > 0:
            print(f"  … 외 {more}건 (리포트 참고)")
    return verdict, ck


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("file")
    ap.add_argument("--profile")
    ap.add_argument("--out")
    ap.add_argument("--json", action="store_true")
    a = ap.parse_args()
    verdict, ck = run(a.file, a.profile, a.out, quiet=a.json)
    if a.json:
        print(json.dumps({"verdict": verdict, "findings": [asdict(f) for f in ck.findings]}, ensure_ascii=False, indent=1))
    sys.exit(1 if verdict == "FAIL" else 0)


if __name__ == "__main__":
    main()
