#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
KCSC(국가건설기준센터) 표준시방서 자동 수집·치환·워드(.docx) 생성기
===================================================================
로컬 PC 폴더의 `시방서_설정.xlsx` 파일을 읽어:
  1. `포함여부`가 'Y'인 건설기준 코드만 추출
  2. KCSC Open API(https://kcsc.re.kr/OpenApi/CodeViewer)로 최신 본문(JSON) 요청
  3. BeautifulSoup 기반 HTML 태그 정리 (문단/목록/표 구조 보존)
  4. 엑셀의 `치환규칙` 시트에 따라 문구 일괄 치환 (일반 문자열 및 정규식 지원)
  5. 완성된 A4 규격 워드 파일(.docx) 자동 생성 (표지 요약표, 목차 계층, 네이티브 표, 쪽번호 포함)

필수 패키지 설치:
    pip install requests openpyxl beautifulsoup4 python-docx

사용 예시:
    python kcsc_spec_builder.py
    python kcsc_spec_builder.py --config 시방서_설정.xlsx --output 완성_시방서.docx
    python kcsc_spec_builder.py --init-template   # 샘플 시방서_설정.xlsx 생성
"""

from __future__ import annotations

import argparse
import dataclasses
import datetime
import base64
import html
import io
import logging
import os
import pathlib
import re
import sys
import time
import urllib.parse
from typing import Any, Dict, List, Optional, Tuple

import openpyxl
import requests
from bs4 import BeautifulSoup, Comment, NavigableString, Tag
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Pt, RGBColor
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# ---------------------------------------------------------------------------
# 로깅 설정
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger("kcsc_spec_builder")

KCSC_API_BASE = "https://kcsc.re.kr/OpenApi/CodeViewer"
KCSC_SITE_BASE = "https://www.kcsc.re.kr"
DEFAULT_CONFIG_NAME = "시방서_설정.xlsx"
TRUTHY_VALUES = {"Y", "YES", "예", "O", "TRUE", "1", "사용", "포함"}


# ---------------------------------------------------------------------------
# 데이터 모델
# ---------------------------------------------------------------------------
@dataclasses.dataclass
class SpecCodeItem:
    doc_type: str       # 예: KCS, KDS, SMCS
    doc_code: str       # 예: 114010
    raw_code: str       # 엑셀 원본 문자열
    title_hint: str = ""
    order: int = 0

    @property
    def display_code(self) -> str:
        if len(self.doc_code) == 6 and self.doc_code.isdigit():
            return f"{self.doc_type} {self.doc_code[:2]} {self.doc_code[2:4]} {self.doc_code[4:]}"
        return f"{self.doc_type} {self.doc_code}"


@dataclasses.dataclass
class ReplaceRule:
    find_text: str
    replace_text: str
    is_regex: bool = False
    case_sensitive: bool = True
    note: str = ""
    hit_count: int = 0


@dataclasses.dataclass
class AppConfig:
    api_key: str = ""
    project_title: str = "공사 시방서"
    output_filename: str = ""
    page_break_per_code: bool = True
    codes: List[SpecCodeItem] = dataclasses.field(default_factory=list)
    rules: List[ReplaceRule] = dataclasses.field(default_factory=list)


@dataclasses.dataclass
class TableCell:
    text: str
    rowspan: int = 1
    colspan: int = 1
    is_header: bool = False


@dataclasses.dataclass
class ContentBlock:
    kind: str           # "paragraph" | "list_item" | "table" | "image"
    text: str = ""
    rows: List[List[TableCell]] = dataclasses.field(default_factory=list)
    image_src: str = ""
    image_alt: str = ""


# ---------------------------------------------------------------------------
# 유틸리티: 코드 정규화 및 헤더 매칭
# ---------------------------------------------------------------------------
def _norm_header(val: Any) -> str:
    if val is None:
        return ""
    return re.sub(r"[\s_\-()\[\]]+", "", str(val)).lower()


def _is_truthy(val: Any) -> bool:
    if val is None:
        return False
    return str(val).strip().upper() in TRUTHY_VALUES


VALID_TYPES = {"KCS", "KDS", "SMCS", "NHCS", "LHCS", "KRACS", "KWCS", "EXCS"}


def parse_kcsc_code(raw_code: Any, type_hint: Any = None) -> Tuple[str, str]:
    """
    'KCS 11 40 10', 'KCS114010', '11 40 10'(type='KCS'), 엑셀 숫자형(114010.0) 등
    다양한 표기를 검증 후 (doc_type, doc_code) 튜플로 정규화합니다.
    """
    if raw_code is None:
        raise ValueError("건설기준 코드가 비어 있습니다.")

    # 엑셀 숫자형 셀이 float(예: 114010.0)로 읽히는 경우 보정
    if isinstance(raw_code, float) and raw_code.is_integer():
        text = str(int(raw_code))
    else:
        text = str(raw_code).strip()

    if not text:
        raise ValueError("건설기준 코드가 비어 있습니다.")

    # ': 2025', '(제목)' 등 부가 표기 및 대괄호 제거
    text = re.split(r"[:：\(（]", text)[0].strip()
    text = text.strip("[]").strip()

    hint = str(type_hint or "").strip().upper()
    match = re.fullmatch(r"([A-Za-z]+)?\s*([0-9\s\-]+)", text)
    if not match:
        raise ValueError(f"잘못된 코드 형식: '{raw_code}'")

    doc_type = match.group(1).upper() if match.group(1) else (hint or "KCS")
    doc_code = re.sub(r"[^0-9]", "", match.group(2))

    if doc_type not in VALID_TYPES:
        raise ValueError(f"지원하지 않는 기준 유형: '{doc_type}' (원본: '{raw_code}')")

    if len(doc_code) != 6:
        raise ValueError(f"코드 길이 오류(6자리 필요): '{doc_code}' (원본: '{raw_code}')")

    return doc_type, doc_code


def _find_header_row(ws, candidates: List[str], max_scan: int = 10) -> Tuple[Optional[int], Dict[str, int]]:
    """시트 상단에서 헤더 행을 자동 탐지하고 {정규화컬럼명: 열인덱스(0-based)} 반환."""
    cand_set = {_norm_header(c) for c in candidates}
    for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=max_scan, values_only=True), start=1):
        norm_map: Dict[str, int] = {}
        for col_idx, cell_val in enumerate(row):
            n = _norm_header(cell_val)
            if n:
                norm_map[n] = col_idx
        if norm_map and (cand_set & set(norm_map.keys())):
            return row_idx, norm_map
    return None, {}


def _pick_col(col_map: Dict[str, int], aliases: List[str]) -> Optional[int]:
    for alias in aliases:
        key = _norm_header(alias)
        if key in col_map:
            return col_map[key]
    return None


# ---------------------------------------------------------------------------
# 1. 엑셀 설정 파일 로더
# ---------------------------------------------------------------------------
def validate_unique_codes(codes: List[SpecCodeItem]) -> None:
    """포함 대상 시방서 코드 목록에 중복이 있는지 일괄 검사합니다."""
    seen: Dict[Tuple[str, str], List[str]] = {}
    duplicates: List[Tuple[str, List[str]]] = []

    for item in codes:
        key = (item.doc_type.strip().upper(), item.doc_code.strip())
        if key in seen:
            if len(seen[key]) == 1:
                duplicates.append((item.display_code, seen[key]))
            seen[key].append(item.raw_code)
        else:
            seen[key] = [item.raw_code]

    if duplicates:
        details = [f"- {disp} (엑셀 원본: {', '.join(repr(r) for r in raws)})" for disp, raws in duplicates]
        raise ValueError("중복 시방서 코드 발견:\n" + "\n".join(details))


def load_config_from_excel(xlsx_path: pathlib.Path) -> AppConfig:
    if not xlsx_path.exists():
        raise FileNotFoundError(
            f"설정 파일을 찾을 수 없습니다: {xlsx_path}\n"
            f"샘플 파일이 필요하면 `--init-template` 옵션으로 생성할 수 있습니다."
        )

    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    cfg = AppConfig()

    sheet_names = wb.sheetnames
    logger.info("엑셀 파일 로드: %s (시트: %s)", xlsx_path.name, ", ".join(sheet_names))

    # --- (A) 시방서 코드 목록 시트 찾기 ---
    code_sheet = None
    for name in sheet_names:
        n = _norm_header(name)
        if any(k in n for k in ["시방서", "코드", "목록", "code", "spec", "list"]):
            code_sheet = wb[name]
            break
    if code_sheet is None:
        code_sheet = wb[sheet_names[0]]

    hdr_row, col_map = _find_header_row(
        code_sheet, ["코드", "시방서코드", "기준코드", "문서번호", "code", "포함여부", "include"]
    )
    if hdr_row is None:
        raise ValueError(f"'{code_sheet.title}' 시트에서 코드 목록 헤더를 찾지 못했습니다.")

    idx_include = _pick_col(col_map, ["포함여부", "포함", "사용여부", "적용여부", "적용", "include", "yn"])
    idx_type = _pick_col(col_map, ["타입", "문서타입", "기준구분", "구분", "종류", "type", "codetype"])
    idx_code = _pick_col(col_map, ["코드", "코드번호", "시방서코드", "기준코드", "문서번호", "code"])
    idx_title = _pick_col(col_map, ["제목", "시방서명", "기준명", "코드명", "명칭", "name", "title"])
    idx_order = _pick_col(col_map, ["순서", "정렬", "번호", "no", "order", "sort"])

    if idx_code is None:
        raise ValueError(f"'{code_sheet.title}' 시트에 '코드' 열이 필요합니다.")
    if idx_include is None:
        raise ValueError(f"'{code_sheet.title}' 시트에 '포함여부' 열이 필요합니다.")

    for r_i, row in enumerate(code_sheet.iter_rows(min_row=hdr_row + 1, values_only=True), start=1):
        if not row or all(v is None or str(v).strip() == "" for v in row):
            continue
        raw_inc = row[idx_include] if idx_include < len(row) else None
        if not _is_truthy(raw_inc):
            continue

        raw_code = row[idx_code] if idx_code < len(row) else None
        if raw_code is None or str(raw_code).strip() == "":
            continue

        raw_type = row[idx_type] if (idx_type is not None and idx_type < len(row)) else None
        raw_title = row[idx_title] if (idx_title is not None and idx_title < len(row)) else ""
        raw_order = row[idx_order] if (idx_order is not None and idx_order < len(row)) else r_i

        try:
            doc_type, doc_code = parse_kcsc_code(raw_code, raw_type)
            order_val = int(raw_order) if str(raw_order).strip().isdigit() else r_i
            cfg.codes.append(
                SpecCodeItem(
                    doc_type=doc_type,
                    doc_code=doc_code,
                    raw_code=str(raw_code).strip(),
                    title_hint=str(raw_title or "").strip(),
                    order=order_val,
                )
            )
        except ValueError as e:
            logger.warning("코드 파싱 건너뜀 (행 %d): %s", hdr_row + r_i, e)

    cfg.codes.sort(key=lambda x: x.order)
    validate_unique_codes(cfg.codes)

    # --- (B) 치환규칙 시트 찾기 ---
    rule_sheet = None
    for name in sheet_names:
        n = _norm_header(name)
        if any(k in n for k in ["치환", "규칙", "변경", "replace", "rule"]):
            rule_sheet = wb[name]
            break
    if rule_sheet is None and len(sheet_names) >= 2:
        rule_sheet = wb[sheet_names[1]]

    if rule_sheet is not None and rule_sheet != code_sheet:
        r_hdr, r_map = _find_header_row(
            rule_sheet, ["찾을문구", "변경전", "원본", "원본단어", "대상문구", "find", "before", "old"]
        )
        if r_hdr is not None:
            idx_find = _pick_col(r_map, ["찾을문구", "변경전", "원본", "원본단어", "대상문구", "find", "before", "old"])
            idx_rep = _pick_col(r_map, ["바꿀문구", "변경후", "변경단어", "치환문구", "치환단어", "수정문구", "replace", "after", "new"])
            idx_use = _pick_col(r_map, ["사용여부", "활성여부", "포함여부", "적용", "use", "active"])
            idx_regex = _pick_col(r_map, ["정규식", "정규표현식", "regex"])
            idx_case = _pick_col(r_map, ["대소문자구분", "대소문자", "casesensitive", "case"])
            idx_note = _pick_col(r_map, ["비고", "설명", "note", "memo"])

            if idx_find is not None and idx_rep is not None:
                for row in rule_sheet.iter_rows(min_row=r_hdr + 1, values_only=True):
                    if not row or all(v is None or str(v).strip() == "" for v in row):
                        continue
                    if idx_use is not None and idx_use < len(row):
                        use_val = row[idx_use]
                        if use_val is not None and str(use_val).strip() != "" and not _is_truthy(use_val):
                            continue

                    find_val = row[idx_find] if idx_find < len(row) else None
                    if find_val is None or str(find_val) == "":
                        continue
                    rep_val = row[idx_rep] if idx_rep < len(row) else ""
                    is_regex = _is_truthy(row[idx_regex]) if (idx_regex is not None and idx_regex < len(row)) else False
                    case_sens = True
                    if idx_case is not None and idx_case < len(row) and row[idx_case] is not None:
                        if str(row[idx_case]).strip() != "":
                            case_sens = _is_truthy(row[idx_case])
                    note_val = str(row[idx_note] or "") if (idx_note is not None and idx_note < len(row)) else ""

                    cfg.rules.append(
                        ReplaceRule(
                            find_text=str(find_val),
                            replace_text="" if rep_val is None else str(rep_val),
                            is_regex=is_regex,
                            case_sensitive=case_sens,
                            note=note_val,
                        )
                    )

    # --- (C) 설정 시트 (선택) ---
    cfg_sheet = None
    for name in sheet_names:
        n = _norm_header(name)
        if any(k in n for k in ["설정", "환경", "api", "setting", "config"]):
            cfg_sheet = wb[name]
            break

    if cfg_sheet is not None and cfg_sheet != code_sheet:
        for row in cfg_sheet.iter_rows(min_row=1, max_col=3, values_only=True):
            if not row or len(row) < 2 or row[0] is None:
                continue
            k = _norm_header(row[0])
            v = "" if row[1] is None else str(row[1]).strip()
            if not v:
                continue
            if k in {"apikey", "key", "인증키", "kcscapikey", "api키"}:
                cfg.api_key = v
            elif k in {"프로젝트명", "공사명", "문서제목", "제목", "projecttitle", "title"}:
                cfg.project_title = v
            elif k in {"출력파일명", "파일명", "저장파일명", "output", "filename"}:
                cfg.output_filename = v
            elif k in {"코드별페이지나눔", "페이지나눔", "pagebreak"}:
                cfg.page_break_per_code = _is_truthy(v)

    return cfg


# ---------------------------------------------------------------------------
# 2. KCSC Open API 클라이언트
# ---------------------------------------------------------------------------
def validate_api_response(data: Any, doc_type: str, doc_code: str) -> Dict[str, Any]:
    """KCSC Open API 응답 구조를 검증하고 대소문자 키를 정규화합니다."""
    code_label = f"{doc_type} {doc_code}"

    # 래퍼(dict) 내부에 list/dict가 감싸져 있는 경우 방어적으로 언래핑
    if isinstance(data, dict) and not any(
        k in data for k in ("list", "List", "message", "Message", "error", "code", "Code", "name", "Name")
    ):
        for wrapper_key in ("data", "result", "items", "body"):
            if wrapper_key in data and isinstance(data[wrapper_key], (list, dict)):
                data = data[wrapper_key]
                break

    if isinstance(data, list):
        dict_items = [d for d in data if isinstance(d, dict)]
        if not dict_items:
            raise RuntimeError(f"빈 API 응답 또는 유효한 객체 없음 ({code_label})")
        if len(dict_items) == 1:
            data = dict_items[0]
        else:
            def _score(d: Dict[str, Any]) -> Tuple[int, int, str]:
                d_code = re.sub(r"[^0-9]", "", str(d.get("code") or d.get("Code") or ""))
                has_list = 1 if (d.get("list") or d.get("List")) else 0
                upd = str(d.get("updateDate") or d.get("UpdateDate") or "")
                return (1 if d_code == doc_code else 0, has_list, upd)

            data = max(dict_items, key=_score)
            logger.warning(
                "다중 응답(%d건) 수신 -> 최적 항목 선택 (%s, 버전: %s)",
                len(dict_items), code_label, data.get("version") or data.get("Version") or "-"
            )

    if not isinstance(data, dict):
        raise RuntimeError(f"API 응답 형식 오류 ({code_label}): {type(data).__name__}")

    error = (
        data.get("message")
        or data.get("Message")
        or data.get("error")
        or data.get("errorMessage")
    )
    if error:
        raise RuntimeError(f"API 오류 ({code_label}): {error}")

    # 대소문자 키 혼용(API 명세 대문자 vs 예시 소문자) 정규화
    for upper_k, lower_k in [
        ("Code", "code"),
        ("Name", "name"),
        ("Version", "version"),
        ("UpdateDate", "updateDate"),
        ("List", "list"),
    ]:
        if lower_k not in data and upper_k in data:
            data[lower_k] = data[upper_k]

    if not str(data.get("name") or "").strip():
        raise RuntimeError(f"시방서명이 누락되었습니다 ({code_label})")

    sections = data.get("list")
    if not isinstance(sections, list) or not sections:
        raise RuntimeError(f"시방서 목차 또는 본문 데이터가 없습니다 ({code_label})")

    normalized_sections: List[Dict[str, Any]] = []
    has_meaningful_content = False
    for idx, section in enumerate(sections, start=1):
        if not isinstance(section, dict):
            raise RuntimeError(f"목차 데이터 형식 오류 ({code_label}, 항목 #{idx})")

        sec = dict(section)
        for upper_k, lower_k in [
            ("Sort", "sort"),
            ("Title", "title"),
            ("Level", "level"),
            ("Label", "label"),
            ("Contents", "contents"),
        ]:
            if lower_k not in sec and upper_k in sec:
                sec[lower_k] = sec[upper_k]

        if str(sec.get("title") or "").strip() or str(sec.get("contents") or "").strip():
            has_meaningful_content = True
        normalized_sections.append(sec)

    if not has_meaningful_content:
        raise RuntimeError(f"시방서 목차와 본문 내용이 모두 비어 있습니다 ({code_label})")

    data["list"] = normalized_sections
    return data


class KCSCClient:
    def __init__(self, api_key: str, timeout: int = 15, delay: float = 0.25):
        self.api_key = api_key.strip()
        self.timeout = timeout
        self.delay = delay
        self.session = requests.Session()
        retry = Retry(
            total=3,
            backoff_factor=0.8,
            status_forcelist=[429, 500, 502, 503, 504],
            allowed_methods=["GET"],
        )
        adapter = HTTPAdapter(max_retries=retry)
        self.session.mount("https://", adapter)
        self.session.mount("http://", adapter)
        self.session.headers.update({
            "Accept": "application/json",
            "User-Agent": "KCSC-Spec-Builder/1.0",
        })

    def fetch_code(self, doc_type: str, doc_code: str) -> Dict[str, Any]:
        url = f"{KCSC_API_BASE}/{doc_type}/{doc_code}"
        params = {"key": self.api_key}
        resp = self.session.get(url, params=params, timeout=self.timeout)
        resp.raise_for_status()

        try:
            raw_data = resp.json()
        except ValueError as exc:
            preview = resp.text[:200].strip().replace("\n", " ")
            raise RuntimeError(
                f"API 응답을 JSON으로 파싱할 수 없습니다 ({doc_type} {doc_code}, HTTP {resp.status_code}): {preview}"
            ) from exc

        item = validate_api_response(raw_data, doc_type, doc_code)
        time.sleep(self.delay)
        return item


# ---------------------------------------------------------------------------
# 3. HTML 태그 정리 및 구조화 파싱
# ---------------------------------------------------------------------------
def _clean_text(text: str) -> str:
    text = html.unescape(text)
    text = text.replace("\xa0", " ").replace("\u200b", "")
    # 연속 공백 정리 (줄바꿈 유지)
    lines = [re.sub(r"[ \t]+", " ", ln).strip() for ln in text.splitlines()]
    cleaned = "\n".join(ln for ln in lines if ln)
    return cleaned.strip()


def _parse_span(val: Any) -> int:
    """rowspan/colspan 속성값을 안전한 양의 정수로 변환합니다."""
    if val is None:
        return 1
    m = re.match(r"^\s*(\d+)", str(val))
    if not m:
        return 1
    n = int(m.group(1))
    return max(n, 1)


def extract_table_cells(table_tag: Tag) -> List[List[TableCell]]:
    """HTML <table>에서 중첩 표를 분리하고 rowspan/colspan을 보존해 TableCell 행렬을 추출합니다."""
    rows: List[List[TableCell]] = []
    for tr in table_tag.find_all("tr"):
        # 중첩된 내부 <table>의 <tr>이 바깥 표로 섞이지 않도록 필터링
        if tr.find_parent("table") != table_tag:
            continue

        cells = tr.find_all(["th", "td"], recursive=False)
        if not cells:
            continue

        row: List[TableCell] = []
        in_thead = tr.find_parent("thead") is not None
        for cell in cells:
            for br in cell.find_all("br"):
                br.replace_with("\n")
            cell_text = _clean_text(cell.get_text(""))
            row.append(
                TableCell(
                    text=cell_text,
                    rowspan=_parse_span(cell.get("rowspan")),
                    colspan=_parse_span(cell.get("colspan")),
                    is_header=(cell.name.lower() == "th" or in_thead),
                )
            )

        if any(c.text for c in row) or any(c.rowspan > 1 or c.colspan > 1 for c in row):
            rows.append(row)

    return rows


_extract_table = extract_table_cells


def parse_html_blocks(raw_html: str) -> List[ContentBlock]:
    """
    KCSC 본문 HTML을 문단, 목록 항목, 네이티브 표 블록으로 분해합니다.
    - HTML 주석(<!-- ... -->), script, style 제거
    - 인라인 태그(<strong>, <span>, <a>, <sup>, <sub> 등)는 문단 분리 없이 하나로 병합
    - <br>은 같은 문단 내 줄바꿈(\n)으로 유지
    - <ul>/<ol> 중첩 목록 및 <li> 내부 표까지 순서대로 추출
    """
    if not raw_html or not str(raw_html).strip():
        return []

    soup = BeautifulSoup(str(raw_html), "html.parser")
    for tag in soup(["script", "style", "noscript"]):
        tag.decompose()
    for comment in soup.find_all(string=lambda t: isinstance(t, Comment)):
        comment.extract()

    blocks: List[ContentBlock] = []
    block_tags = {
        "table", "ul", "ol", "p", "div", "section", "article",
        "blockquote", "figure", "figcaption", "img",
        "h1", "h2", "h3", "h4", "h5", "h6", "hr",
    }

    def _inline_text(tag: Tag) -> str:
        for br in tag.find_all("br"):
            br.replace_with("\n")
        return tag.get_text("")

    def _flush_inline(buf: List[str]):
        if not buf:
            return
        merged = _clean_text("".join(buf))
        buf.clear()
        if merged:
            blocks.append(ContentBlock(kind="paragraph", text=merged))

    def _parse_list(list_tag: Tag, depth: int = 0):
        is_ordered = list_tag.name.lower() == "ol"
        indent = "  " * depth
        for idx, li in enumerate(list_tag.find_all("li", recursive=False), start=1):
            li_buf: List[str] = []

            def _flush_li():
                if not li_buf:
                    return
                txt = _clean_text("".join(li_buf))
                li_buf.clear()
                if not txt:
                    return
                if re.match(r"^([•\-\*]|\(\d+\)|\d+\.)", txt):
                    prefix = ""
                elif is_ordered:
                    prefix = f"{idx}. "
                else:
                    prefix = "• "
                blocks.append(ContentBlock(kind="list_item", text=f"{indent}{prefix}{txt}"))

            for child in li.children:
                if isinstance(child, Comment):
                    continue
                if isinstance(child, NavigableString):
                    li_buf.append(str(child))
                    continue
                if not isinstance(child, Tag):
                    continue
                cname = child.name.lower()
                if cname == "br":
                    li_buf.append("\n")
                elif cname in {"ul", "ol"}:
                    _flush_li()
                    _parse_list(child, depth + 1)
                elif cname == "img":
                    _flush_li()
                    img_src = str(child.get("src") or "").strip()
                    img_alt = str(child.get("alt") or "").strip()
                    if img_src:
                        blocks.append(ContentBlock(kind="image", image_src=img_src, image_alt=img_alt))
                elif cname == "table":
                    _flush_li()
                    rows = _extract_table(child)
                    if rows:
                        blocks.append(ContentBlock(kind="table", rows=rows))
                elif cname in block_tags:
                    _flush_li()
                    walk(child)
                else:
                    li_buf.append(_inline_text(child))
            _flush_li()

    def walk(node: Tag):
        inline_buf: List[str] = []

        for child in node.children:
            if isinstance(child, Comment):
                continue
            if isinstance(child, NavigableString):
                inline_buf.append(str(child))
                continue
            if not isinstance(child, Tag):
                continue

            name = child.name.lower()
            if name == "br":
                inline_buf.append("\n")
            elif name == "hr":
                _flush_inline(inline_buf)
            elif name == "img":
                _flush_inline(inline_buf)
                img_src = str(child.get("src") or "").strip()
                img_alt = str(child.get("alt") or "").strip()
                if img_src:
                    blocks.append(ContentBlock(kind="image", image_src=img_src, image_alt=img_alt))
            elif name == "table":
                _flush_inline(inline_buf)
                rows = _extract_table(child)
                if rows:
                    blocks.append(ContentBlock(kind="table", rows=rows))
            elif name in {"ul", "ol"}:
                _flush_inline(inline_buf)
                _parse_list(child, depth=0)
            elif name in {"p", "div", "section", "article", "blockquote", "figure", "figcaption", "h1", "h2", "h3", "h4", "h5", "h6"}:
                _flush_inline(inline_buf)
                # 내부에 블록 요소가 있으면 재귀 탐색, 인라인 요소만 있으면 단일 문단으로 추출
                if child.find(list(block_tags)):
                    walk(child)
                else:
                    txt = _clean_text(_inline_text(child))
                    if txt:
                        blocks.append(ContentBlock(kind="paragraph", text=txt))
            else:
                # <strong>, <span>, <a>, <sup>, <sub> 등 인라인 태그는 버퍼에 이어 붙임
                inline_buf.append(_inline_text(child))

        _flush_inline(inline_buf)

    root = soup.body if soup.body else soup
    walk(root)
    return blocks


# ---------------------------------------------------------------------------
# 4. 문구 일괄 치환 엔진
# ---------------------------------------------------------------------------
def compile_rules(rules: List[ReplaceRule]) -> List[Tuple[ReplaceRule, re.Pattern]]:
    """
    치환 규칙 목록을 일괄 검증·컴파일합니다.
    - 빈 검색어 건너뛰기
    - 정규식 패턴 문법 오류 및 치환식(\1 등 그룹 참조) 오류를 한 번에 수집해 보고
    - 빈 문자열에 매칭되는 정규식(.* 등) 차단
    """
    compiled: List[Tuple[ReplaceRule, re.Pattern]] = []
    errors: List[str] = []

    for idx, rule in enumerate(rules, start=1):
        if rule.find_text is None or str(rule.find_text) == "":
            continue

        flags = 0 if rule.case_sensitive else re.IGNORECASE
        pattern_str = rule.find_text if rule.is_regex else re.escape(rule.find_text)

        try:
            pat = re.compile(pattern_str, flags)
        except re.error as exc:
            errors.append(f"{idx}번 규칙('{rule.find_text}') 정규식 오류: {exc}")
            continue

        if rule.is_regex:
            # 치환 문자열 내 잘못된 그룹 참조(\9 등) 사전 검증
            try:
                pat.sub(rule.replace_text, "")
            except re.error as exc:
                errors.append(f"{idx}번 규칙('{rule.find_text}' -> '{rule.replace_text}') 치환식 오류: {exc}")
                continue

            if pat.search(""):
                errors.append(f"{idx}번 규칙('{rule.find_text}')은 빈 문자열에도 매칭되어 무한/과다 치환을 유발할 수 있습니다")
                continue

        compiled.append((rule, pat))

    if errors:
        raise ValueError("치환 규칙 검증 실패:\n" + "\n".join(errors))

    return compiled


class ReplacementEngine:
    def __init__(self, rules: List[ReplaceRule]):
        self.rules = rules
        self.patterns = compile_rules(rules)

    def apply(self, text: str) -> str:
        if not text or not self.patterns:
            return text
        result = text
        had_deletion = False
        for rule, pattern in self.patterns:
            rep = "" if rule.replace_text is None else str(rule.replace_text)
            if rule.is_regex:
                # 정규식 규칙은 역참조(\1, \g<1>) 지원
                result, count = pattern.subn(rep, result)
            else:
                # 일반 문자열 규칙은 바꿀 문구의 백슬래시(\)가 이스케이프로 해석되지 않도록 함수로 반환
                result, count = pattern.subn(lambda _m, r=rep: r, result)
            if count > 0 and rep == "":
                had_deletion = True
            rule.hit_count += count

        # 문구 삭제로 인해 남은 중복 공백 및 행 앞뒤 공백 정리 (목록 들여쓰기 2칸은 유지)
        if had_deletion:
            cleaned_lines = []
            for ln in result.splitlines():
                lead = ln[:len(ln) - len(ln.lstrip(" "))]
                body = re.sub(r" {2,}", " ", ln.lstrip(" ")).rstrip(" ")
                if body:
                    cleaned_lines.append((lead if lead == "  " else "") + body)
            result = "\n".join(cleaned_lines)
        return result


# ---------------------------------------------------------------------------
# 5. 워드(.docx) 문서 빌더
# ---------------------------------------------------------------------------
def _set_run_font(run, font_name: str = "맑은 고딕", size_pt: float = 10.0,
                  bold: bool = False, color_rgb: Optional[RGBColor] = None):
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.bold = bold
    if color_rgb:
        run.font.color.rgb = color_rgb
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:eastAsia"), font_name)
    rFonts.set(qn("w:ascii"), font_name)
    rFonts.set(qn("w:hAnsi"), font_name)


def _set_cell_shading(cell, fill_hex: str):
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading)


def _set_cell_margins(cell, top=80, bottom=80, left=120, right=120):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for side, val in [("top", top), ("bottom", bottom), ("left", left), ("right", right)]:
        node = OxmlElement(f"w:{side}")
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)


def _set_table_borders(table, color: str = "BFBFBF", sz: str = "4"):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:bottom w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:left w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:right w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:insideH w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:insideV w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f"</w:tblBorders>"
    )
    tblPr.append(borders)


def set_repeat_table_header(row, enabled: bool = True):
    """페이지가 넘어갈 때 표의 머리글 행이 자동으로 반복되도록 w:tblHeader를 안전하게 설정합니다."""
    if row is None or not hasattr(row, "_tr"):
        return
    trPr = row._tr.get_or_add_trPr()
    existing = trPr.find(qn("w:tblHeader"))
    if not enabled:
        if existing is not None:
            trPr.remove(existing)
        return
    if existing is None:
        tbl_header = OxmlElement("w:tblHeader")
        tbl_header.set(qn("w:val"), "true")
        trPr.append(tbl_header)
    else:
        existing.set(qn("w:val"), "true")


def _add_page_number_field(paragraph):
    run_pre = paragraph.add_run("- ")
    _set_run_font(run_pre, size_pt=9, color_rgb=RGBColor(0x66, 0x66, 0x66))

    for field_name in ["PAGE", " / ", "NUMPAGES"]:
        if field_name == " / ":
            r = paragraph.add_run(" / ")
            _set_run_font(r, size_pt=9, color_rgb=RGBColor(0x66, 0x66, 0x66))
            continue
        r_begin = paragraph.add_run()
        _set_run_font(r_begin, size_pt=9, color_rgb=RGBColor(0x66, 0x66, 0x66))
        fld_begin = OxmlElement("w:fldChar")
        fld_begin.set(qn("w:fldCharType"), "begin")
        r_begin._r.append(fld_begin)

        r_instr = paragraph.add_run()
        _set_run_font(r_instr, size_pt=9, color_rgb=RGBColor(0x66, 0x66, 0x66))
        instr = OxmlElement("w:instrText")
        instr.set(qn("xml:space"), "preserve")
        instr.text = f" {field_name} "
        r_instr._r.append(instr)

        r_end = paragraph.add_run()
        _set_run_font(r_end, size_pt=9, color_rgb=RGBColor(0x66, 0x66, 0x66))
        fld_end = OxmlElement("w:fldChar")
        fld_end.set(qn("w:fldCharType"), "end")
        r_end._r.append(fld_end)

    run_post = paragraph.add_run(" -")
    _set_run_font(run_post, size_pt=9, color_rgb=RGBColor(0x66, 0x66, 0x66))


class SpecDocxBuilder:
    PRIMARY_HEX = "1F497D"
    PRIMARY_RGB = RGBColor(0x1F, 0x49, 0x7D)
    DARK_RGB = RGBColor(0x22, 0x22, 0x22)
    MUTED_RGB = RGBColor(0x55, 0x55, 0x55)

    def __init__(self, project_title: str, image_session: Optional[requests.Session] = None):
        self.doc = Document()
        self.project_title = project_title
        self._image_session = image_session
        self._image_cache: Dict[str, Optional[bytes]] = {}
        self._configure_page_and_styles()

    def _configure_page_and_styles(self):
        section = self.doc.sections[0]
        section.page_width = Cm(21.0)
        section.page_height = Cm(29.7)
        section.top_margin = Cm(2.0)
        section.bottom_margin = Cm(2.0)
        section.left_margin = Cm(2.2)
        section.right_margin = Cm(2.2)
        section.different_first_page_header_footer = True

        # 기본 Normal 스타일 한글 폰트 적용
        normal = self.doc.styles["Normal"]
        normal.font.name = "맑은 고딕"
        normal.font.size = Pt(10)
        normal.font.color.rgb = self.DARK_RGB
        rPr = normal._element.get_or_add_rPr()
        rFonts = OxmlElement("w:rFonts")
        rFonts.set(qn("w:eastAsia"), "맑은 고딕")
        rFonts.set(qn("w:ascii"), "맑은 고딕")
        rFonts.set(qn("w:hAnsi"), "맑은 고딕")
        rPr.insert(0, rFonts)

        # 머리글 / 바닥글
        hdr_p = section.header.paragraphs[0]
        hdr_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        hdr_run = hdr_p.add_run(self.project_title)
        _set_run_font(hdr_run, size_pt=8.5, color_rgb=self.MUTED_RGB)

        ftr_p = section.footer.paragraphs[0]
        ftr_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _add_page_number_field(ftr_p)

    def add_cover_summary(self, fetched_specs: List[Dict[str, Any]], replacer: ReplacementEngine):
        p_title = self.doc.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_title.paragraph_format.space_before = Pt(36)
        p_title.paragraph_format.space_after = Pt(8)
        r_title = p_title.add_run(replacer.apply(self.project_title))
        _set_run_font(r_title, size_pt=22, bold=True, color_rgb=self.PRIMARY_RGB)

        p_sub = self.doc.add_paragraph()
        p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_sub.paragraph_format.space_after = Pt(20)
        today_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
        r_sub = p_sub.add_run(f"국가건설기준센터(KCSC) 최신 표준시방서 기반 자동 생성 | 생성일시: {today_str}")
        _set_run_font(r_sub, size_pt=9.5, color_rgb=self.MUTED_RGB)

        p_sec = self.doc.add_paragraph()
        p_sec.paragraph_format.space_before = Pt(10)
        p_sec.paragraph_format.space_after = Pt(6)
        r_sec = p_sec.add_run("수록 건설기준 목록")
        _set_run_font(r_sec, size_pt=12, bold=True, color_rgb=self.PRIMARY_RGB)

        headers = ["순번", "기준코드", "시방서명", "버전", "최종수정일"]
        tbl = self.doc.add_table(rows=1 + len(fetched_specs), cols=5)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        _set_table_borders(tbl)
        set_repeat_table_header(tbl.rows[0])

        col_widths = [Cm(1.4), Cm(3.4), Cm(7.6), Cm(1.8), Cm(2.4)]
        for c_idx, h_text in enumerate(headers):
            cell = tbl.cell(0, c_idx)
            cell.width = col_widths[c_idx]
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            _set_cell_shading(cell, self.PRIMARY_HEX)
            _set_cell_margins(cell)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(h_text)
            _set_run_font(r, size_pt=9.5, bold=True, color_rgb=RGBColor(0xFF, 0xFF, 0xFF))

        for r_idx, spec in enumerate(fetched_specs, start=1):
            meta: SpecCodeItem = spec["meta"]
            data = spec["data"]
            update_dt = str(data.get("updateDate") or "")[:10]
            row_vals = [
                str(r_idx),
                meta.display_code,
                replacer.apply(str(data.get("name") or meta.title_hint)),
                str(data.get("version") or "-"),
                update_dt or "-",
            ]
            for c_idx, val in enumerate(row_vals):
                cell = tbl.cell(r_idx, c_idx)
                cell.width = col_widths[c_idx]
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                _set_cell_margins(cell)
                if r_idx % 2 == 0:
                    _set_cell_shading(cell, "F5F7FA")
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT if c_idx == 2 else WD_ALIGN_PARAGRAPH.CENTER
                r = p.add_run(val)
                _set_run_font(r, size_pt=9.5)

    def add_toc(self, max_level: int = 4, page_break_before: bool = True):
        """Word 자동 목차(TOC) 필드를 삽입하고 문서 열기 시 자동 업데이트를 활성화합니다."""
        if page_break_before:
            self.doc.add_page_break()

        p_hdr = self.doc.add_paragraph()
        p_hdr.paragraph_format.space_before = Pt(4)
        p_hdr.paragraph_format.space_after = Pt(10)
        p_hdr.paragraph_format.keep_with_next = True
        r_hdr = p_hdr.add_run("목  차")
        _set_run_font(r_hdr, size_pt=14, bold=True, color_rgb=self.PRIMARY_RGB)

        # 문서 열기 시 Word가 목차 등 필드를 자동 갱신하도록 settings.xml에 updateFields 설정
        settings_el = self.doc.settings.element
        if settings_el.find(qn("w:updateFields")) is None:
            upd = OxmlElement("w:updateFields")
            upd.set(qn("w:val"), "true")
            settings_el.append(upd)

        p = self.doc.add_paragraph()
        p.paragraph_format.space_after = Pt(6)

        # OpenXML 규격상 fldChar / instrText / w:t는 각각 별도 <w:r>로 분리해야 안전하게 동작
        r_begin = p.add_run()
        fld_begin = OxmlElement("w:fldChar")
        fld_begin.set(qn("w:fldCharType"), "begin")
        fld_begin.set(qn("w:dirty"), "true")
        r_begin._r.append(fld_begin)

        r_instr = p.add_run()
        instr_text = OxmlElement("w:instrText")
        instr_text.set(qn("xml:space"), "preserve")
        instr_text.text = f' TOC \\o "1-{max_level}" \\h \\z \\u '
        r_instr._r.append(instr_text)

        r_sep = p.add_run()
        fld_sep = OxmlElement("w:fldChar")
        fld_sep.set(qn("w:fldCharType"), "separate")
        r_sep._r.append(fld_sep)

        r_msg = p.add_run("목차를 표시하려면 Word에서 [F9] 또는 '필드 업데이트'를 실행하세요.")
        _set_run_font(r_msg, size_pt=9.5, color_rgb=self.MUTED_RGB)

        r_end = p.add_run()
        fld_end = OxmlElement("w:fldChar")
        fld_end.set(qn("w:fldCharType"), "end")
        r_end._r.append(fld_end)

    def add_spec_document(self, meta: SpecCodeItem, data: Dict[str, Any],
                          replacer: ReplacementEngine, page_break: bool = True):
        if page_break:
            self.doc.add_page_break()

        name = replacer.apply(str(data.get("name") or meta.title_hint))
        version = str(data.get("version") or "")
        update_dt = str(data.get("updateDate") or "")[:10]

        # 코드 대제목 (Heading 1)
        h1 = self.doc.add_heading(level=1)
        h1.paragraph_format.space_before = Pt(6)
        h1.paragraph_format.space_after = Pt(4)
        h1.paragraph_format.keep_with_next = True
        r_h1 = h1.add_run(f"[{meta.display_code}] {name}")
        _set_run_font(r_h1, size_pt=15, bold=True, color_rgb=self.PRIMARY_RGB)

        # 메타 정보 라인
        meta_parts = [f"분류: {meta.doc_type}"]
        if version:
            meta_parts.append(f"버전: {version}")
        if update_dt:
            meta_parts.append(f"최종수정일: {update_dt}")
        p_meta = self.doc.add_paragraph()
        p_meta.paragraph_format.space_after = Pt(10)
        r_meta = p_meta.add_run(" | ".join(meta_parts))
        _set_run_font(r_meta, size_pt=8.5, color_rgb=self.MUTED_RGB)

        sections = data.get("list") or []
        sections = sorted(sections, key=lambda s: int(s.get("sort") or 0))

        for sec in sections:
            raw_title = str(sec.get("title") or "").strip()
            level = int(sec.get("level") or 1)
            contents = str(sec.get("contents") or "")

            if raw_title:
                h_level = min(max(level + 1, 2), 4)
                h_p = self.doc.add_heading(level=h_level)
                h_p.paragraph_format.space_before = Pt(10 if level == 1 else 6)
                h_p.paragraph_format.space_after = Pt(3)
                h_p.paragraph_format.keep_with_next = True
                h_size = {2: 12.5, 3: 11.0, 4: 10.5}.get(h_level, 10.5)
                h_color = self.PRIMARY_RGB if h_level == 2 else self.DARK_RGB
                r_t = h_p.add_run(replacer.apply(raw_title))
                _set_run_font(r_t, size_pt=h_size, bold=True, color_rgb=h_color)

            blocks = parse_html_blocks(contents)
            for blk in blocks:
                if blk.kind == "table":
                    self._add_native_table(blk.rows, replacer)
                elif blk.kind == "image":
                    self._add_image(blk.image_src, blk.image_alt, replacer)
                elif blk.kind == "list_item":
                    txt = replacer.apply(blk.text)
                    if not txt.strip() or re.fullmatch(r"\s*(?:[•\-\*]|\d+\.)\s*", txt):
                        continue
                    p = self.doc.add_paragraph()
                    p.paragraph_format.left_indent = Cm(0.6)
                    p.paragraph_format.space_after = Pt(3)
                    p.paragraph_format.line_spacing = 1.25
                    r = p.add_run(txt)
                    _set_run_font(r, size_pt=10)
                else:
                    txt = replacer.apply(blk.text)
                    if not txt.strip():
                        continue
                    # 목차 제목이 본문 첫 줄에 중복 출력되는 경우 방지
                    if raw_title and txt == replacer.apply(raw_title):
                        continue
                    p = self.doc.add_paragraph()
                    p.paragraph_format.space_after = Pt(4)
                    p.paragraph_format.line_spacing = 1.25
                    r = p.add_run(txt)
                    _set_run_font(r, size_pt=10)

    def _add_native_table(self, rows: List[List[Any]], replacer: ReplacementEngine):
        if not rows:
            return

        # List[List[str]] 호환 및 TableCell 정규화
        norm_rows: List[List[TableCell]] = []
        for row in rows:
            norm_row: List[TableCell] = []
            for item in row:
                if isinstance(item, TableCell):
                    norm_row.append(item)
                else:
                    norm_row.append(TableCell(text=str(item), is_header=(len(norm_rows) == 0)))
            if norm_row:
                norm_rows.append(norm_row)

        if not norm_rows:
            return

        # 2D 점유 그리드 계산: rowspan/colspan으로 밀리는 열·행 좌표를 정확히 산출
        occupied: set = set()
        placements: List[Tuple[int, int, int, int, TableCell]] = []
        total_rows = len(norm_rows)
        total_cols = 0

        for r_idx, row in enumerate(norm_rows):
            c_idx = 0
            for cell_obj in row:
                while (r_idx, c_idx) in occupied:
                    c_idx += 1
                r_span = max(cell_obj.rowspan, 1)
                c_span = max(cell_obj.colspan, 1)
                end_row = r_idx + r_span - 1
                end_col = c_idx + c_span - 1

                for rr in range(r_idx, end_row + 1):
                    for cc in range(c_idx, end_col + 1):
                        occupied.add((rr, cc))

                total_rows = max(total_rows, end_row + 1)
                total_cols = max(total_cols, end_col + 1)
                placements.append((r_idx, c_idx, end_row, end_col, cell_obj))
                c_idx += c_span

        if total_cols == 0:
            return

        tbl = self.doc.add_table(rows=total_rows, cols=total_cols)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        _set_table_borders(tbl)

        for r_idx in range(total_rows):
            trPr = tbl.rows[r_idx]._tr.get_or_add_trPr()
            trPr.append(OxmlElement("w:cantSplit"))
            is_hdr_row = (r_idx == 0) or (
                r_idx < len(norm_rows) and all(c.is_header for c in norm_rows[r_idx])
            )
            if is_hdr_row:
                set_repeat_table_header(tbl.rows[r_idx])

        for start_row, start_col, end_row, end_col, cell_obj in placements:
            end_row = min(end_row, total_rows - 1)
            end_col = min(end_col, total_cols - 1)

            top_left = tbl.cell(start_row, start_col)
            top_left.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            _set_cell_margins(top_left, top=60, bottom=60, left=100, right=100)
            if cell_obj.is_header or start_row == 0:
                _set_cell_shading(top_left, "EAEFF5")

            val = replacer.apply(cell_obj.text)
            p = top_left.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(val)
            _set_run_font(r, size_pt=9.0, bold=(cell_obj.is_header or start_row == 0))

            if end_row > start_row or end_col > start_col:
                merged_cell = tbl.cell(start_row, start_col).merge(tbl.cell(end_row, end_col))
                merged_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        spacer = self.doc.add_paragraph()
        spacer.paragraph_format.space_after = Pt(4)

    def _fetch_image_bytes(self, src: str) -> Optional[bytes]:
        src = src.strip()
        if not src:
            return None
        if src in self._image_cache:
            return self._image_cache[src]

        raw_bytes: Optional[bytes] = None
        if src.startswith("data:image/"):
            try:
                _, b64_part = src.split(",", 1)
                raw_bytes = base64.b64decode(b64_part)
            except Exception as exc:
                logger.warning("Base64 이미지 디코딩 실패: %s", exc)
                self._image_cache[src] = None
                return None
        else:
            if src.startswith("//"):
                url = "https:" + src
            elif src.startswith(("http://", "https://")):
                url = src
            else:
                url = urllib.parse.urljoin(KCSC_SITE_BASE + "/", src.lstrip("/"))

            sess = self._image_session or requests.Session()
            try:
                resp = sess.get(url, timeout=15)
                resp.raise_for_status()
                raw_bytes = resp.content
            except Exception as exc:
                logger.warning("이미지 다운로드 실패 (%s): %s", url, exc)
                self._image_cache[src] = None
                return None

        # Pillow가 있으면 WebP/RGBA 등 호환성 문제를 PNG로 정규화
        try:
            from PIL import Image as PILImage  # optional
            with PILImage.open(io.BytesIO(raw_bytes)) as im:
                fmt = (im.format or "").upper()
                if fmt not in {"PNG", "JPEG", "JPG", "BMP", "GIF", "TIFF"}:
                    buf = io.BytesIO()
                    im.convert("RGBA" if "A" in im.getbands() else "RGB").save(buf, format="PNG")
                    raw_bytes = buf.getvalue()
        except ImportError:
            pass
        except Exception as exc:
            logger.warning("이미지 포맷 확인 실패 (%s): %s", src[:60], exc)
            self._image_cache[src] = None
            return None

        self._image_cache[src] = raw_bytes
        return raw_bytes

    def _add_image(self, src: str, alt: str, replacer: ReplacementEngine):
        img_bytes = self._fetch_image_bytes(src)
        if not img_bytes:
            fallback_label = replacer.apply(alt) if alt else src[:80]
            p = self.doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_after = Pt(6)
            r = p.add_run(f"[이미지 로드 실패: {fallback_label}]")
            r.italic = True
            _set_run_font(r, size_pt=9.0, color_rgb=self.MUTED_RGB)
            return

        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(6)
        run = p.add_run()

        max_w_cm = 15.5
        max_h_cm = 20.0
        width_arg = Cm(max_w_cm)
        height_arg = None

        try:
            from PIL import Image as PILImage
            with PILImage.open(io.BytesIO(img_bytes)) as im:
                px_w, px_h = im.size
                if px_w > 0 and px_h > 0:
                    dpi = im.info.get("dpi", (96, 96))
                    dpi_x = dpi[0] if isinstance(dpi, tuple) and dpi[0] > 0 else 96
                    native_w_cm = (px_w / dpi_x) * 2.54
                    native_h_cm = (px_h / dpi_x) * 2.54
                    scale = min(1.0, max_w_cm / native_w_cm, max_h_cm / native_h_cm)
                    width_arg = Cm(native_w_cm * scale)
                    height_arg = Cm(native_h_cm * scale)
        except Exception:
            pass

        try:
            if height_arg is not None:
                run.add_picture(io.BytesIO(img_bytes), width=width_arg, height=height_arg)
            else:
                run.add_picture(io.BytesIO(img_bytes), width=width_arg)
        except Exception as exc:
            logger.warning("워드 이미지 삽입 실패 (%s): %s", src[:60], exc)
            r = p.add_run(f"[이미지 삽입 불가: {alt or src[:60]}]")
            _set_run_font(r, size_pt=9.0, color_rgb=self.MUTED_RGB)

    def save(self, out_path: pathlib.Path) -> pathlib.Path:
        out_path.parent.mkdir(parents=True, exist_ok=True)
        try:
            self.doc.save(str(out_path))
            return out_path
        except PermissionError:
            # Word에서 기존 파일이 열려 있는 경우 타임스탬프를 붙여 안전 저장
            stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            fallback = out_path.with_name(f"{out_path.stem}_{stamp}{out_path.suffix}")
            self.doc.save(str(fallback))
            logger.warning("대상 파일이 사용 중이어서 새 이름으로 저장했습니다: %s", fallback)
            return fallback


# ---------------------------------------------------------------------------
# 6. 샘플 시방서_설정.xlsx 생성기 (--init-template)
# ---------------------------------------------------------------------------
def create_template_excel(target_path: pathlib.Path):
    wb = openpyxl.Workbook()

    # 시트 1: 시방서목록
    ws1 = wb.active
    ws1.title = "시방서목록"
    ws1.append(["순서", "포함여부", "타입", "코드", "시방서명(참고)"])
    ws1.append([1, "Y", "KCS", "114010", "파형강판 암거"])
    ws1.append([2, "Y", "KCS", "142010", "일반콘크리트"])
    ws1.append([3, "N", "KCS", "112015", "터파기 (미포함 예시)"])

    # 시트 2: 치환규칙
    ws2 = wb.create_sheet("치환규칙")
    ws2.append(["사용여부", "찾을문구", "바꿀문구", "정규식", "비고"])
    ws2.append(["Y", "공사감독자", "건설사업관리기술인", "N", "용어 통일"])
    ws2.append(["Y", "수급인", "시공자", "N", "용어 통일"])

    # 시트 3: 설정
    ws3 = wb.create_sheet("설정")
    ws3.append(["항목", "값", "설명"])
    ws3.append(["API_KEY", "", "KCSC Open API에서 발급받은 인증키 입력"])
    ws3.append(["프로젝트명", "OO건설공사 표준시방서", "워드 문서 표지 및 머리글 제목"])
    ws3.append(["출력파일명", "완성_시방서.docx", "비워두면 프로젝트명_날짜.docx로 저장"])
    ws3.append(["코드별페이지나눔", "Y", "각 시방서 코드마다 새 페이지에서 시작"])

    wb.save(str(target_path))
    logger.info("샘플 설정 파일을 생성했습니다: %s", target_path.resolve())


# ---------------------------------------------------------------------------
# 메인 실행 파이프라인
# ---------------------------------------------------------------------------
def resolve_api_key(cli_key: Optional[str], cfg_key: Optional[str]) -> str:
    """
    우선순위(CLI 옵션 > 환경변수 KCSC_API_KEY > 엑셀 설정 시트 > 대화형 입력)로
    인증키를 안전하게 결정하고 불필요한 공백·따옴표·중괄호를 제거합니다.
    """
    def _clean(val: Optional[str]) -> str:
        return str(val or "").strip().strip("\"'{}")

    candidates = [
        ("CLI 옵션(--api-key)", _clean(cli_key)),
        ("환경변수(KCSC_API_KEY)", _clean(os.environ.get("KCSC_API_KEY"))),
        ("엑셀 설정 시트", _clean(cfg_key)),
    ]
    for source, key in candidates:
        if key:
            logger.info("KCSC API 인증키 로드 완료 (출처: %s)", source)
            return key

    if sys.stdin.isatty():
        entered = _clean(input("KCSC Open API 인증키를 입력하세요: "))
        if entered:
            logger.info("KCSC API 인증키 로드 완료 (출처: 직접 입력)")
            return entered

    return ""


def validate_fetch_results(
    codes: List[SpecCodeItem],
    fetched_specs: List[Dict[str, Any]],
    failed_codes: List[str],
    allow_partial: bool = False,
) -> None:
    """수집 대상 건수와 성공/실패 건수 및 코드 집합 일치 여부를 검증합니다."""
    expected = len(codes)
    success = len(fetched_specs)
    failed = len(failed_codes)

    logger.info("전체 %d건 / 성공 %d건 / 실패 %d건", expected, success, failed)

    if expected == 0:
        raise RuntimeError("수집 대상 시방서 코드가 0건입니다.")

    if success + failed != expected:
        raise RuntimeError(
            f"수집 건수 불일치: 전체 {expected}건 != 성공 {success}건 + 실패 {failed}건"
        )

    if failed > 0:
        detail_msg = "\n".join(failed_codes)
        logger.error("일부 기준 수집에 실패했습니다 (%d/%d건):\n%s", failed, expected, detail_msg)
        if not allow_partial:
            raise RuntimeError(
                f"완전한 시방서 생성이 불가능합니다 (실패 {failed}건):\n{detail_msg}"
            )
        logger.warning("--allow-partial 옵션이 활성화되어 성공한 %d건만으로 문서를 생성합니다.", success)

    if not allow_partial:
        if success != expected:
            raise RuntimeError(
                f"수집 건수 검증에 실패했습니다 (예상 {expected}건, 성공 {success}건)."
            )
        expected_keys = {(c.doc_type, c.doc_code) for c in codes}
        fetched_keys = {(s["meta"].doc_type, s["meta"].doc_code) for s in fetched_specs}
        missing_keys = expected_keys - fetched_keys
        if missing_keys:
            missing_str = ", ".join(f"{t} {c}" for t, c in sorted(missing_keys))
            raise RuntimeError(f"누락된 시방서 코드가 있습니다: {missing_str}")

    if success == 0:
        raise RuntimeError("수신에 성공한 시방서가 없어 워드 파일을 생성하지 않습니다.")


def main():
    parser = argparse.ArgumentParser(
        description="시방서_설정.xlsx 기반으로 KCSC API 본문을 수집·치환하여 워드(.docx)로 생성합니다."
    )
    parser.add_argument(
        "--config", "-c",
        default=DEFAULT_CONFIG_NAME,
        help=f"설정 엑셀 파일 경로 (기본값: {DEFAULT_CONFIG_NAME})",
    )
    parser.add_argument(
        "--output", "-o",
        default="",
        help="출력 워드(.docx) 파일 경로",
    )
    parser.add_argument(
        "--api-key", "-k",
        default="",
        help="KCSC Open API 인증키 (환경변수 KCSC_API_KEY 또는 엑셀 설정 시트 사용 가능)",
    )
    parser.add_argument(
        "--init-template",
        action="store_true",
        help="현재 폴더에 샘플 시방서_설정.xlsx 파일을 생성하고 종료합니다.",
    )
    parser.add_argument(
        "--allow-partial",
        action="store_true",
        help="일부 코드 조회에 실패해도 성공한 코드만으로 워드 파일을 생성합니다.",
    )
    args = parser.parse_args()

    if args.init_template:
        create_template_excel(pathlib.Path(args.config))
        return

    config_path = pathlib.Path(args.config)
    if not config_path.exists():
        # 스크립트와 같은 폴더에서도 탐색
        script_dir_cfg = pathlib.Path(__file__).resolve().parent / args.config
        if script_dir_cfg.exists():
            config_path = script_dir_cfg

    cfg = load_config_from_excel(config_path)

    if not cfg.codes:
        logger.error("포함여부가 'Y'인 시방서 코드가 없습니다. 엑셀 파일을 확인해주세요.")
        sys.exit(1)

    api_key = resolve_api_key(args.api_key, cfg.api_key)
    if not api_key:
        logger.error(
            "KCSC API 인증키가 필요합니다. "
            "엑셀 '설정' 시트의 API_KEY, 환경변수 KCSC_API_KEY, 또는 --api-key 옵션을 사용하세요."
        )
        sys.exit(1)

    logger.info("대상 시방서 %d건 | 치환규칙 %d건 로드 완료", len(cfg.codes), len(cfg.rules))

    client = KCSCClient(api_key=api_key)
    fetched_specs: List[Dict[str, Any]] = []
    failed_codes: List[str] = []

    for idx, item in enumerate(cfg.codes, start=1):
        logger.info("[%d/%d] %s (%s) 요청 중...", idx, len(cfg.codes), item.display_code, item.doc_code)
        try:
            data = client.fetch_code(item.doc_type, item.doc_code)
            sec_cnt = len(data.get("list") or [])
            logger.info("  -> 수신 완료: %s (목차 %d건)", data.get("name", ""), sec_cnt)
            fetched_specs.append({"meta": item, "data": data})
        except Exception as exc:
            logger.error("  -> 조회 실패 (%s): %s", item.display_code, exc)
            failed_codes.append(f"{item.display_code}: {exc}")

    validate_fetch_results(
        cfg.codes,
        fetched_specs,
        failed_codes,
        allow_partial=args.allow_partial,
    )

    replacer = ReplacementEngine(cfg.rules)
    builder = SpecDocxBuilder(project_title=cfg.project_title, image_session=client.session)
    builder.add_cover_summary(fetched_specs, replacer)
    builder.add_toc()

    for spec in fetched_specs:
        builder.add_spec_document(
            meta=spec["meta"],
            data=spec["data"],
            replacer=replacer,
            page_break=cfg.page_break_per_code,
        )

    if args.output:
        out_path = pathlib.Path(args.output)
    elif cfg.output_filename:
        out_path = config_path.parent / cfg.output_filename
    else:
        today = datetime.datetime.now().strftime("%Y%m%d")
        safe_title = re.sub(r'[\\/:*?"<>|]+', "_", cfg.project_title).strip() or "시방서"
        out_path = config_path.parent / f"{safe_title}_{today}.docx"

    saved_path = builder.save(out_path)

    total_hits = sum(r.hit_count for r in cfg.rules)
    logger.info("=" * 60)
    logger.info("워드 파일 생성 완료: %s", saved_path.resolve())
    logger.info("수집 성공: %d건 | 실패: %d건 | 문구 치환 총 %d회 적용",
                len(fetched_specs), len(failed_codes), total_hits)
    for r in cfg.rules:
        if r.hit_count > 0:
            logger.info("  - '%s' -> '%s': %d회", r.find_text, r.replace_text, r.hit_count)
    if failed_codes:
        logger.warning("실패한 코드 목록:\n  %s", "\n  ".join(failed_codes))


if __name__ == "__main__":
    main()