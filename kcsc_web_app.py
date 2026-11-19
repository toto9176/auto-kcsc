#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
KCSC 시방서 자동화 - 프로페셔널 다크모드 PC·모바일 웹/PWA 앱
================================================================
같은 폴더의 kcsc_spec_builder.py 엔진을 사용하여:
  - PC 브라우저에서 '앱 설치'로 데스크톱 프로그램처럼 실행
  - 같은 Wi-Fi의 스마트폰 브라우저에서 '홈 화면에 추가'로 모바일 앱처럼 실행
  - 시방서_설정.xlsx 자동 로드 및 업로드, 코드 Y/N 토글, 치환규칙 편집,
    실시간 진행률·로그 확인 및 완성 워드(.docx) 다운로드 지원

실행 방법:
    python kcsc_web_app.py
    python kcsc_web_app.py --port 8501 --no-browser
"""

from __future__ import annotations

import argparse
import base64
import datetime
import json
import os
import pathlib
import re
import socket
import sys
import threading
import time
import urllib.parse
import webbrowser
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from typing import Any, Dict, List, Optional

import openpyxl

BASE_DIR = pathlib.Path(__file__).resolve().parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

import kcsc_spec_builder as ksb  # noqa: E402

OUTPUT_DIR = BASE_DIR / "output_docs"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


def _get_lan_ip() -> str:
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as s:
            s.connect(("8.8.8.8", 80))
            return s.getsockname()[0]
    except OSError:
        return "127.0.0.1"


# ---------------------------------------------------------------------------
# 앱 전역 상태 관리 (스레드 안전)
# ---------------------------------------------------------------------------
class AppState:
    def __init__(self):
        self.lock = threading.Lock()
        self.codes: List[Dict[str, Any]] = []
        self.rules: List[Dict[str, Any]] = []
        self.settings: Dict[str, Any] = {
            "api_key": os.environ.get("KCSC_API_KEY", "").strip(),
            "project_title": "공사 시방서",
            "output_filename": "완성_시방서.docx",
            "page_break_per_code": True,
        }
        self.excel_name: str = ""
        self.excel_path: Optional[pathlib.Path] = None
        self.last_excel_path: Optional[pathlib.Path] = None
        self.running: bool = False
        self.progress_done: int = 0
        self.progress_total: int = 0
        self.logs: List[str] = []
        self.last_error: str = ""
        self.last_output_path: Optional[pathlib.Path] = None
        self.last_summary: Dict[str, Any] = {}

    def add_log(self, msg: str):
        stamp = datetime.datetime.now().strftime("%H:%M:%S")
        with self.lock:
            self.logs.append(f"[{stamp}] {msg}")
            if len(self.logs) > 300:
                self.logs = self.logs[-300:]

    def load_excel_path(self, xlsx_path: pathlib.Path, display_name: str = ""):
        wb = openpyxl.load_workbook(xlsx_path, data_only=True)
        sheet_names = wb.sheetnames

        # 1. 목록 시트 전체 행(Y/N 포함) 로드
        code_sheet = None
        for name in sheet_names:
            n = ksb._norm_header(name)
            if any(k in n for k in ["시방서", "코드", "목록", "code", "spec", "list"]):
                code_sheet = wb[name]
                break
        if code_sheet is None:
            code_sheet = wb[sheet_names[0]]

        hdr_row, col_map = ksb._find_header_row(
            code_sheet, ["코드", "시방서코드", "기준코드", "문서번호", "code", "포함여부", "include"]
        )
        if hdr_row is None:
            raise ValueError(f"'{code_sheet.title}' 시트에서 헤더를 찾지 못했습니다.")

        idx_inc = ksb._pick_col(col_map, ["포함여부", "포함", "사용여부", "적용여부", "적용", "include", "yn"])
        idx_type = ksb._pick_col(col_map, ["타입", "문서타입", "기준구분", "구분", "종류", "type", "codetype"])
        idx_code = ksb._pick_col(col_map, ["코드", "코드번호", "시방서코드", "기준코드", "문서번호", "code"])
        idx_title = ksb._pick_col(col_map, ["제목", "시방서명", "기준명", "코드명", "명칭", "name", "title"])
        idx_cat = ksb._pick_col(col_map, ["중공종", "공종", "분류", "category"])

        parsed_codes: List[Dict[str, Any]] = []
        for r_i, row in enumerate(code_sheet.iter_rows(min_row=hdr_row + 1, values_only=True), start=1):
            if not row or all(v is None or str(v).strip() == "" for v in row):
                continue
            raw_code = row[idx_code] if (idx_code is not None and idx_code < len(row)) else None
            if raw_code is None or str(raw_code).strip() == "":
                continue
            raw_type = row[idx_type] if (idx_type is not None and idx_type < len(row)) else None
            raw_inc = row[idx_inc] if (idx_inc is not None and idx_inc < len(row)) else "N"
            raw_title = row[idx_title] if (idx_title is not None and idx_title < len(row)) else ""
            raw_cat = row[idx_cat] if (idx_cat is not None and idx_cat < len(row)) else ""
            try:
                doc_type, doc_code = ksb.parse_kcsc_code(raw_code, raw_type)
                parsed_codes.append({
                    "id": r_i,
                    "row_num": hdr_row + r_i,
                    "include": ksb._is_truthy(raw_inc),
                    "doc_type": doc_type,
                    "doc_code": doc_code,
                    "display_code": f"{doc_type} {doc_code[:2]} {doc_code[2:4]} {doc_code[4:]}",
                    "raw_code": str(raw_code).strip(),
                    "title": str(raw_title or "").strip(),
                    "category": str(raw_cat or "").strip(),
                })
            except ValueError:
                continue

        # 2. 치환규칙 + 설정 시트는 기존 ksb 로더 활용
        cfg = ksb.load_config_from_excel(xlsx_path)
        parsed_rules = [
            {
                "find_text": r.find_text,
                "replace_text": r.replace_text,
                "is_regex": r.is_regex,
                "case_sensitive": r.case_sensitive,
                "note": r.note,
            }
            for r in cfg.rules
        ]

        with self.lock:
            self.codes = parsed_codes
            self.rules = parsed_rules
            if cfg.api_key:
                self.settings["api_key"] = cfg.api_key
            if cfg.project_title:
                self.settings["project_title"] = cfg.project_title
            if cfg.output_filename:
                self.settings["output_filename"] = cfg.output_filename
            self.settings["page_break_per_code"] = cfg.page_break_per_code
            self.excel_name = display_name or xlsx_path.name
            self.excel_path = xlsx_path
            self.last_excel_path = xlsx_path

        self.add_log(
            f"엑셀 로드 완료: {self.excel_name} "
            f"(전체 {len(parsed_codes)}건 중 포함 {sum(1 for c in parsed_codes if c['include'])}건, "
            f"치환규칙 {len(parsed_rules)}건)"
        )

    def save_excel_changes(self) -> Optional[pathlib.Path]:
        """앱에서 수정한 목록(Y/N), 치환규칙, 설정을 원본(또는 새) 시방서_설정.xlsx에 반영합니다."""
        with self.lock:
            target_path = self.excel_path or (BASE_DIR / "시방서_설정.xlsx")
            codes_copy = list(self.codes)
            rules_copy = list(self.rules)
            settings_copy = dict(self.settings)

        if target_path.exists():
            wb = openpyxl.load_workbook(target_path)
        else:
            wb = openpyxl.Workbook()
            ws0 = wb.active
            ws0.title = "목록"
            ws0.append(["포함여부", "기준구분", "기준코드", "시방서명", "중 공 종"])

        sheet_names = wb.sheetnames

        # --- (1) 목록 시트 포함여부(Y/N) 업데이트 ---
        code_sheet = None
        for name in sheet_names:
            n = ksb._norm_header(name)
            if any(k in n for k in ["시방서", "코드", "목록", "code", "spec", "list"]):
                code_sheet = wb[name]
                break
        if code_sheet is None:
            code_sheet = wb[sheet_names[0]]

        hdr_row, col_map = ksb._find_header_row(
            code_sheet, ["코드", "시방서코드", "기준코드", "문서번호", "code", "포함여부", "include"]
        )
        if hdr_row is not None:
            idx_inc = ksb._pick_col(col_map, ["포함여부", "포함", "사용여부", "적용여부", "적용", "include", "yn"])
            idx_code = ksb._pick_col(col_map, ["코드", "코드번호", "시방서코드", "기준코드", "문서번호", "code"])
            idx_type = ksb._pick_col(col_map, ["타입", "문서타입", "기준구분", "구분", "종류", "type", "codetype"])

            if idx_inc is not None:
                by_row = {c.get("row_num"): c for c in codes_copy if c.get("row_num")}
                by_key = {(c["doc_type"], c["doc_code"]): c for c in codes_copy}

                for r_idx in range(hdr_row + 1, code_sheet.max_row + 1):
                    matched = by_row.get(r_idx)
                    if matched is None and idx_code is not None:
                        raw_c = code_sheet.cell(row=r_idx, column=idx_code + 1).value
                        raw_t = code_sheet.cell(row=r_idx, column=idx_type + 1).value if idx_type is not None else None
                        if raw_c is not None and str(raw_c).strip():
                            try:
                                dt, dc = ksb.parse_kcsc_code(raw_c, raw_t)
                                matched = by_key.get((dt, dc))
                            except ValueError:
                                matched = None
                    if matched is not None:
                        code_sheet.cell(row=r_idx, column=idx_inc + 1).value = "Y" if matched.get("include") else "N"

        # --- (2) 치환규칙 시트 업데이트 ---
        rule_sheet = None
        for name in sheet_names:
            n = ksb._norm_header(name)
            if any(k in n for k in ["치환", "규칙", "변경", "replace", "rule"]):
                rule_sheet = wb[name]
                break
        if rule_sheet is None:
            rule_sheet = wb.create_sheet("치환규칙")
            rule_sheet.append(["사용여부", "원본단어", "변경단어", "정규식", "비고"])
            r_hdr = 1
            r_map = {"사용여부": 0, "원본단어": 1, "변경단어": 2, "정규식": 3, "비고": 4}
        else:
            r_hdr, r_map = ksb._find_header_row(
                rule_sheet, ["찾을문구", "변경전", "원본", "원본단어", "대상문구", "find", "before", "old"]
            )
            if r_hdr is None:
                r_hdr = 1
                rule_sheet.cell(row=1, column=1, value="원본단어")
                rule_sheet.cell(row=1, column=2, value="변경단어")
                r_map = {"원본단어": 0, "변경단어": 1}

        idx_find = ksb._pick_col(r_map, ["찾을문구", "변경전", "원본", "원본단어", "대상문구", "find", "before", "old"])
        idx_rep = ksb._pick_col(r_map, ["바꿀문구", "변경후", "변경단어", "치환문구", "치환단어", "수정문구", "replace", "after", "new"])
        idx_use = ksb._pick_col(r_map, ["사용여부", "활성여부", "포함여부", "적용", "use", "active"])
        idx_regex = ksb._pick_col(r_map, ["정규식", "정규표현식", "regex"])
        idx_note = ksb._pick_col(r_map, ["비고", "설명", "note", "memo"])

        # 기존 규칙 행 초기화 후 현재 앱 상태 기록
        max_rule_cols = max(rule_sheet.max_column, (max(r_map.values()) + 1) if r_map else 2)
        for rr in range(r_hdr + 1, rule_sheet.max_row + 1):
            for cc in range(1, max_rule_cols + 1):
                rule_sheet.cell(row=rr, column=cc).value = None

        has_regex_rule = any(bool(r.get("is_regex")) for r in rules_copy)
        if has_regex_rule and idx_regex is None:
            idx_regex = max_rule_cols
            rule_sheet.cell(row=r_hdr, column=idx_regex + 1, value="정규식")

        for offset, rule in enumerate(rules_copy, start=r_hdr + 1):
            if not str(rule.get("find_text") or "").strip():
                continue
            if idx_use is not None:
                rule_sheet.cell(row=offset, column=idx_use + 1).value = "Y"
            if idx_find is not None:
                rule_sheet.cell(row=offset, column=idx_find + 1).value = str(rule.get("find_text") or "")
            if idx_rep is not None:
                rep_str = str(rule.get("replace_text") or "")
                rule_sheet.cell(row=offset, column=idx_rep + 1).value = rep_str if rep_str != "" else None
            if idx_regex is not None:
                rule_sheet.cell(row=offset, column=idx_regex + 1).value = "Y" if rule.get("is_regex") else "N"
            if idx_note is not None:
                rule_sheet.cell(row=offset, column=idx_note + 1).value = str(rule.get("note") or "")

        # --- (3) 설정 시트 업데이트 ---
        cfg_sheet = None
        for name in wb.sheetnames:
            n = ksb._norm_header(name)
            if any(k in n for k in ["설정", "환경", "api", "setting", "config"]):
                cfg_sheet = wb[name]
                break
        if cfg_sheet is None:
            cfg_sheet = wb.create_sheet("설정")
            cfg_sheet.append(["항목", "값", "설명"])

        cfg_values = {
            "apikey": settings_copy.get("api_key", ""),
            "프로젝트명": settings_copy.get("project_title", "공사 시방서"),
            "출력파일명": settings_copy.get("output_filename", "완성_시방서.docx"),
            "코드별페이지나눔": "Y" if settings_copy.get("page_break_per_code", True) else "N",
        }
        updated_keys = set()
        for rr in range(1, cfg_sheet.max_row + 1):
            k_val = cfg_sheet.cell(row=rr, column=1).value
            nk = ksb._norm_header(k_val)
            if nk in {"apikey", "key", "인증키", "kcscapikey", "api키"}:
                cfg_sheet.cell(row=rr, column=2).value = cfg_values["apikey"]
                updated_keys.add("apikey")
            elif nk in {"프로젝트명", "공사명", "문서제목", "제목", "projecttitle", "title"}:
                cfg_sheet.cell(row=rr, column=2).value = cfg_values["프로젝트명"]
                updated_keys.add("프로젝트명")
            elif nk in {"출력파일명", "파일명", "저장파일명", "output", "filename"}:
                cfg_sheet.cell(row=rr, column=2).value = cfg_values["출력파일명"]
                updated_keys.add("출력파일명")
            elif nk in {"코드별페이지나눔", "페이지나눔", "pagebreak"}:
                cfg_sheet.cell(row=rr, column=2).value = cfg_values["코드별페이지나눔"]
                updated_keys.add("코드별페이지나눔")

        missing_defs = [
            ("apikey", "API_KEY", "국가건설기준센터 Open API 인증키"),
            ("프로젝트명", "프로젝트명", "워드 표지 및 머리글 제목"),
            ("출력파일명", "출력파일명", "저장할 워드 파일명"),
            ("코드별페이지나눔", "코드별페이지나눔", "각 시방서 코드 시작 시 새 페이지 구분 (Y/N)"),
        ]
        for norm_k, label, desc in missing_defs:
            if norm_k not in updated_keys:
                cfg_sheet.append([label, cfg_values[norm_k], desc])

        try:
            wb.save(str(target_path))
            saved_xlsx = target_path
        except PermissionError:
            stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            fallback = OUTPUT_DIR / f"{target_path.stem}_{stamp}.xlsx"
            wb.save(str(fallback))
            saved_xlsx = fallback
            self.add_log(f"원본 엑셀이 열려 있어 새 파일에 설정을 저장했습니다: {fallback.name}")

        with self.lock:
            self.last_excel_path = saved_xlsx
        return saved_xlsx

    def snapshot(self) -> Dict[str, Any]:
        with self.lock:
            return {
                "excel_name": self.excel_name,
                "codes": self.codes,
                "rules": self.rules,
                "settings": {
                    "api_key_set": bool(self.settings.get("api_key")),
                    "api_key": self.settings.get("api_key", ""),
                    "project_title": self.settings.get("project_title", "공사 시방서"),
                    "output_filename": self.settings.get("output_filename", "완성_시방서.docx"),
                    "page_break_per_code": self.settings.get("page_break_per_code", True),
                },
                "running": self.running,
                "progress_done": self.progress_done,
                "progress_total": self.progress_total,
                "logs": list(self.logs),
                "last_error": self.last_error,
                "download_ready": bool(self.last_output_path and self.last_output_path.exists()),
                "download_filename": self.last_output_path.name if self.last_output_path else "",
                "excel_download_ready": bool(self.last_excel_path and self.last_excel_path.exists()),
                "excel_download_filename": self.last_excel_path.name if self.last_excel_path else "",
                "last_summary": self.last_summary,
                "lan_url": f"http://{_get_lan_ip()}:{SERVER_PORT}",
            }


STATE = AppState()
SERVER_PORT = 8501


# ---------------------------------------------------------------------------
# 백그라운드 생성 워커
# ---------------------------------------------------------------------------
def run_generation_job(payload: Dict[str, Any]):
    try:
        with STATE.lock:
            if STATE.running:
                return
            STATE.running = True
            STATE.last_error = ""
            STATE.last_summary = {}
            STATE.progress_done = 0

        # UI에서 편집한 상태 반영
        selected_codes_raw = payload.get("codes")
        if isinstance(selected_codes_raw, list):
            with STATE.lock:
                STATE.codes = selected_codes_raw

        rules_raw = payload.get("rules")
        if isinstance(rules_raw, list):
            with STATE.lock:
                STATE.rules = rules_raw

        settings_raw = payload.get("settings") or {}
        with STATE.lock:
            if "api_key" in settings_raw:
                STATE.settings["api_key"] = str(settings_raw["api_key"]).strip().strip("\"'{} ")
            if "project_title" in settings_raw and str(settings_raw["project_title"]).strip():
                STATE.settings["project_title"] = str(settings_raw["project_title"]).strip()
            if "output_filename" in settings_raw:
                STATE.settings["output_filename"] = str(settings_raw["output_filename"]).strip()
            if "page_break_per_code" in settings_raw:
                STATE.settings["page_break_per_code"] = bool(settings_raw["page_break_per_code"])

            active_codes = [
                ksb.SpecCodeItem(
                    doc_type=c["doc_type"],
                    doc_code=c["doc_code"],
                    raw_code=c.get("raw_code", c["doc_code"]),
                    title_hint=c.get("title", ""),
                    order=i + 1,
                )
                for i, c in enumerate(STATE.codes)
                if c.get("include")
            ]
            active_rules = [
                ksb.ReplaceRule(
                    find_text=str(r.get("find_text", "")),
                    replace_text=str(r.get("replace_text") or ""),
                    is_regex=bool(r.get("is_regex", False)),
                    case_sensitive=bool(r.get("case_sensitive", True)),
                    note=str(r.get("note", "")),
                )
                for r in STATE.rules
                if str(r.get("find_text", "")).strip()
            ]
            api_key = STATE.settings["api_key"] or os.environ.get("KCSC_API_KEY", "").strip()
            project_title = STATE.settings["project_title"]
            output_filename = STATE.settings["output_filename"]
            page_break = STATE.settings["page_break_per_code"]
            STATE.progress_total = len(active_codes)

        if not active_codes:
            raise ValueError("포함(Y)으로 선택된 시방서 코드가 없습니다.")
        if not api_key:
            raise ValueError("KCSC Open API 인증키를 입력해주세요.")

        ksb.validate_unique_codes(active_codes)
        saved_cfg_xlsx = STATE.save_excel_changes()
        if saved_cfg_xlsx:
            STATE.add_log(f"수정된 설정 엑셀 자동 저장: {saved_cfg_xlsx.name}")
        replacer = ksb.ReplacementEngine(active_rules)
        allow_partial = bool(payload.get("allow_partial", False))

        STATE.add_log(f"생성 시작: 대상 {len(active_codes)}건, 치환규칙 {len(active_rules)}건")
        client = ksb.KCSCClient(api_key=api_key)
        fetched_specs: List[Dict[str, Any]] = []
        failed_codes: List[str] = []

        for idx, item in enumerate(active_codes, start=1):
            STATE.add_log(f"[{idx}/{len(active_codes)}] {item.display_code} ({item.title_hint or item.doc_code}) 요청 중...")
            try:
                data = client.fetch_code(item.doc_type, item.doc_code)
                fetched_specs.append({"meta": item, "data": data})
                STATE.add_log(f"  -> 수신 완료: {data.get('name', '')} (목차 {len(data.get('list') or [])}건)")
            except Exception as exc:
                failed_codes.append(f"{item.display_code}: {exc}")
                STATE.add_log(f"  -> 조회 실패 ({item.display_code}): {exc}")
            finally:
                with STATE.lock:
                    STATE.progress_done = idx

        ksb.validate_fetch_results(active_codes, fetched_specs, failed_codes, allow_partial=allow_partial)

        STATE.add_log("HTML 본문 정리 및 워드(.docx) 조립 중...")
        builder = ksb.SpecDocxBuilder(project_title=project_title, image_session=client.session)
        builder.add_cover_summary(fetched_specs, replacer)
        builder.add_toc()
        for spec in fetched_specs:
            builder.add_spec_document(
                meta=spec["meta"],
                data=spec["data"],
                replacer=replacer,
                page_break=page_break,
            )

        if not output_filename:
            today = datetime.datetime.now().strftime("%Y%m%d")
            safe_title = re.sub(r'[\\/:*?"<>|]+', "_", project_title).strip() or "시방서"
            output_filename = f"{safe_title}_{today}.docx"
        if not output_filename.lower().endswith(".docx"):
            output_filename += ".docx"

        saved_path = builder.save(OUTPUT_DIR / pathlib.Path(output_filename).name)
        total_hits = sum(r.hit_count for r in active_rules)
        with STATE.lock:
            STATE.last_output_path = saved_path
            STATE.last_summary = {
                "success": len(fetched_specs),
                "failed": len(failed_codes),
                "total_hits": total_hits,
                "filename": saved_path.name,
            }
        STATE.add_log(f"완료! 파일 저장됨: {saved_path.name} (성공 {len(fetched_specs)}건, 치환 {total_hits}회)")
    except Exception as exc:
        with STATE.lock:
            STATE.last_error = str(exc)
        STATE.add_log(f"오류 발생: {exc}")
    finally:
        with STATE.lock:
            STATE.running = False


# ---------------------------------------------------------------------------
# 프로페셔널 다크모드 PWA HTML/CSS/JS
# ---------------------------------------------------------------------------
MANIFEST_JSON = json.dumps({
    "name": "KCSC 시방서 자동화",
    "short_name": "KCSC시방서",
    "start_url": "/",
    "display": "standalone",
    "background_color": "#12161F",
    "theme_color": "#00D1FF",
    "icons": [
        {
            "src": "/icon.svg",
            "sizes": "any",
            "type": "image/svg+xml",
            "purpose": "any maskable",
        }
    ],
}, ensure_ascii=False)

ICON_SVG = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 128 128">
  <rect width="128" height="128" rx="28" fill="#12161F"/>
  <rect x="10" y="10" width="108" height="108" rx="22" fill="#1A2233" stroke="#00D1FF" stroke-width="4"/>
  <path d="M36 34h40l16 16v44a6 6 0 0 1-6 6H36a6 6 0 0 1-6-6V40a6 6 0 0 1 6-6z" fill="#0F172A" stroke="#38BDF8" stroke-width="4"/>
  <path d="M76 34v16h16" fill="none" stroke="#38BDF8" stroke-width="4"/>
  <path d="M44 66h40M44 78h40M44 90h26" stroke="#00D1FF" stroke-width="5" stroke-linecap="round"/>
</svg>"""

HTML_PAGE = r"""<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0, viewport-fit=cover" />
  <meta name="theme-color" content="#12161F" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <meta name="apple-mobile-web-app-status-bar-style" content="black-translucent" />
  <link rel="manifest" href="/manifest.webmanifest" />
  <link rel="icon" href="/icon.svg" type="image/svg+xml" />
  <title>KCSC 시방서 자동화 (Pro Dark)</title>
  <style>
    :root {
      --bg: #12161F;
      --surface: #1A2130;
      --surface-2: #222B3D;
      --border: #2D384F;
      --text: #E8EEF8;
      --muted: #93A3BD;
      --cyan: #00D1FF;
      --cyan-dim: rgba(0, 209, 255, 0.14);
      --green: #10B981;
      --green-dim: rgba(16, 185, 129, 0.16);
      --red: #F87171;
      --amber: #FBBF24;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      font-family: 'Malgun Gothic', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
      background: var(--bg);
      color: var(--text);
      line-height: 1.5;
    }
    .topbar {
      position: sticky; top: 0; z-index: 20;
      background: rgba(18, 22, 31, 0.92);
      backdrop-filter: blur(10px);
      border-bottom: 1px solid var(--border);
      padding: 12px 20px;
      display: flex; align-items: center; justify-content: space-between; gap: 12px;
      flex-wrap: wrap;
    }
    .brand { display: flex; align-items: center; gap: 10px; font-weight: 700; font-size: 16px; }
    .dot { width: 10px; height: 10px; border-radius: 50%; background: var(--cyan); box-shadow: 0 0 10px var(--cyan); }
    .badge {
      font-size: 12px; padding: 4px 10px; border-radius: 999px;
      background: var(--green-dim); color: var(--green); border: 1px solid rgba(16,185,129,0.35);
      display: inline-flex; align-items: center; gap: 6px;
    }
    .badge.warn { background: rgba(251,191,36,0.12); color: var(--amber); border-color: rgba(251,191,36,0.35); }
    .container { max-width: 1280px; margin: 0 auto; padding: 18px; display: grid; gap: 16px; }
    .grid-2 { display: grid; grid-template-columns: 1.15fr 0.85fr; gap: 16px; }
    @media (max-width: 900px) { .grid-2 { grid-template-columns: 1fr; } }
    .card {
      background: var(--surface); border: 1px solid var(--border); border-radius: 14px;
      padding: 16px; box-shadow: 0 8px 24px rgba(0,0,0,0.28);
    }
    .card-title {
      font-size: 14px; font-weight: 700; color: var(--text);
      margin: 0 0 12px; display: flex; align-items: center; justify-content: space-between; gap: 8px;
    }
    .sub { font-size: 12px; color: var(--muted); font-weight: 400; }
    .dropzone {
      border: 1.5px dashed #3B4B6B; border-radius: 10px; padding: 16px; text-align: center;
      background: var(--surface-2); cursor: pointer; transition: border-color 0.2s;
    }
    .dropzone:hover, .dropzone.drag { border-color: var(--cyan); background: rgba(0,209,255,0.06); }
    .row { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin-bottom: 10px; }
    @media (max-width: 600px) { .row { grid-template-columns: 1fr; } }
    label { display: block; font-size: 12px; color: var(--muted); margin-bottom: 4px; }
    input[type="text"], input[type="password"] {
      width: 100%; padding: 9px 11px; border-radius: 8px; border: 1px solid var(--border);
      background: #0F141D; color: var(--text); font-size: 13px; outline: none;
    }
    input:focus { border-color: var(--cyan); }
    .btn {
      border: none; border-radius: 9px; padding: 10px 15px; font-size: 13px; font-weight: 700;
      cursor: pointer; display: inline-flex; align-items: center; justify-content: center; gap: 6px;
      transition: opacity 0.15s, transform 0.05s;
    }
    .btn:active { transform: translateY(1px); }
    .btn:disabled { opacity: 0.45; cursor: not-allowed; }
    .btn-cyan { background: linear-gradient(135deg, #00D1FF, #0095FF); color: #08101C; }
    .btn-ghost { background: var(--surface-2); color: var(--text); border: 1px solid var(--border); }
    .btn-green { background: var(--green); color: #052016; }
    .progress-wrap { background: #0E131B; border-radius: 999px; height: 10px; overflow: hidden; border: 1px solid var(--border); }
    .progress-bar { height: 100%; width: 0%; background: linear-gradient(90deg, #00D1FF, #38BDF8); transition: width 0.25s; }
    .log-box {
      background: #0B0F17; border: 1px solid var(--border); border-radius: 10px;
      padding: 12px; height: 210px; overflow-y: auto;
      font-family: ui-monospace, SFMono-Regular, Consolas, monospace; font-size: 12px; color: #A9C1DF;
      white-space: pre-wrap;
    }
    .table-wrap { max-height: 330px; overflow: auto; border: 1px solid var(--border); border-radius: 10px; }
    table { width: 100%; border-collapse: collapse; font-size: 13px; }
    th, td { padding: 8px 10px; border-bottom: 1px solid var(--border); text-align: left; }
    th { position: sticky; top: 0; background: var(--surface-2); color: var(--muted); font-size: 12px; z-index: 2; }
    .toggle {
      appearance: none; width: 42px; height: 22px; border-radius: 999px;
      background: #334155; position: relative; cursor: pointer; transition: background 0.2s;
    }
    .toggle:checked { background: var(--green); }
    .toggle::after {
      content: ""; position: absolute; top: 3px; left: 3px; width: 16px; height: 16px;
      border-radius: 50%; background: #fff; transition: transform 0.2s;
    }
    .toggle:checked::after { transform: translateX(20px); }
    .actions { display: flex; flex-wrap: wrap; gap: 10px; align-items: center; }
    .error-banner {
      display: none; background: rgba(248, 113, 113, 0.14); border: 1px solid rgba(248, 113, 113, 0.45);
      color: #FECACA; padding: 10px 12px; border-radius: 9px; font-size: 13px; margin-top: 10px;
    }
  </style>
</head>
<body>
  <header class="topbar">
    <div class="brand">
      <span class="dot"></span>
      <span>KCSC 시방서 자동화</span>
      <span class="sub">Pro Dark</span>
    </div>
    <div class="actions">
      <span id="apiBadge" class="badge warn">API 키 미입력</span>
      <span id="lanInfo" class="sub"></span>
    </div>
  </header>

  <main class="container">
    <div class="grid-2">
      <!-- 좌측: 설정 & 엑셀 -->
      <section class="card">
        <div class="card-title">
          <span>1. 엑셀 설정 &amp; 프로젝트 환경</span>
          <span id="excelName" class="sub">불러온 파일 없음</span>
        </div>
        <div id="dropzone" class="dropzone">
          <input id="fileInput" type="file" accept=".xlsx" hidden />
          <div style="font-weight:700; margin-bottom:4px;">시방서_설정.xlsx 드래그 또는 클릭하여 업로드</div>
          <div class="sub">목록 · 치환규칙 · 설정 시트를 자동으로 인식합니다.</div>
        </div>

        <div class="row" style="margin-top:12px;">
          <div>
            <label>KCSC Open API 인증키</label>
            <input id="apiKeyInput" type="password" placeholder="발급받은 API Key 입력" />
          </div>
          <div>
            <label>프로젝트명 (워드 표지/머리글)</label>
            <input id="projectTitleInput" type="text" value="공사 시방서" />
          </div>
        </div>
        <div class="row">
          <div>
            <label>출력 워드 파일명</label>
            <input id="outputFilenameInput" type="text" value="완성_시방서.docx" />
          </div>
          <div style="display:flex; align-items:flex-end; gap:14px; padding-bottom:6px;">
            <label style="display:flex; align-items:center; gap:8px; margin:0; cursor:pointer;">
              <input id="pageBreakToggle" type="checkbox" class="toggle" checked />
              <span>코드별 페이지 나눔</span>
            </label>
            <label style="display:flex; align-items:center; gap:8px; margin:0; cursor:pointer;">
              <input id="allowPartialToggle" type="checkbox" class="toggle" />
              <span>실패 시 부분 생성</span>
            </label>
          </div>
        </div>
      </section>

      <!-- 우측: 진행률 & 실시간 로그 -->
      <section class="card">
        <div class="card-title">
          <span>2. 작업 진행률 &amp; 실시간 실행 로그</span>
          <span id="progressText" style="color:var(--cyan); font-weight:700;">0% (0 / 0)</span>
        </div>
        <div class="progress-wrap">
          <div id="progressBar" class="progress-bar"></div>
        </div>
        <div class="actions" style="margin:12px 0;">
          <button id="runBtn" class="btn btn-cyan" style="flex:1;">시방서 자동 생성 실행</button>
          <a id="downloadBtn" class="btn btn-green" href="/api/download" style="display:none; text-decoration:none;">.docx 다운로드</a>
          <a id="downloadExcelBtn" class="btn btn-ghost" href="/api/download-excel" style="display:none; text-decoration:none;">수정된 .xlsx 다운로드</a>
        </div>
        <div id="errorBanner" class="error-banner"></div>
        <div id="logBox" class="log-box">대기 중...</div>
      </section>
    </div>

    <div class="grid-2">
      <!-- 기준 코드 목록 -->
      <section class="card">
        <div class="card-title">
          <span>3. 건설기준 코드 목록 (<span id="codeCount">0</span>건 포함)</span>
          <div class="actions">
            <input id="codeSearch" type="text" placeholder="코드·명칭 검색..." style="width:160px; padding:5px 9px;" />
            <button id="selectAllBtn" class="btn btn-ghost" style="padding:5px 10px;">전체 Y</button>
            <button id="clearAllBtn" class="btn btn-ghost" style="padding:5px 10px;">전체 N</button>
          </div>
        </div>
        <div class="table-wrap">
          <table>
            <thead>
              <tr><th>포함</th><th>기준코드</th><th>시방서명</th><th>공종</th></tr>
            </thead>
            <tbody id="codeTableBody"></tbody>
          </table>
        </div>
      </section>

      <!-- 치환 규칙 -->
      <section class="card">
        <div class="card-title">
          <span>4. 문구 일괄 치환 규칙 (<span id="ruleCount">0</span>건 · 변경 문구 공란 시 삭제)</span>
          <button id="addRuleBtn" class="btn btn-ghost" style="padding:5px 10px;">+ 규칙 추가</button>
        </div>
        <div class="table-wrap">
          <table>
            <thead>
              <tr><th>원본 문구</th><th>변경 문구</th><th>정규식</th><th>삭제</th></tr>
            </thead>
            <tbody id="ruleTableBody"></tbody>
          </table>
        </div>
      </section>
    </div>
  </main>

  <script>
    let appState = { codes: [], rules: [], settings: {} };
    let pollTimer = null;

    const $ = (id) => document.getElementById(id);

    async function fetchState() {
      const res = await fetch('/api/state');
      const data = await res.json();
      applyServerState(data, true);
    }

    function applyServerState(data, syncInputs = false) {
      if (syncInputs) {
        appState.codes = data.codes || [];
        appState.rules = data.rules || [];
        appState.settings = data.settings || {};
        $('apiKeyInput').value = appState.settings.api_key || '';
        $('projectTitleInput').value = appState.settings.project_title || '공사 시방서';
        $('outputFilenameInput').value = appState.settings.output_filename || '완성_시방서.docx';
        $('pageBreakToggle').checked = appState.settings.page_break_per_code !== false;
        renderCodes();
        renderRules();
      }
      $('excelName').textContent = data.excel_name || '불러온 파일 없음';
      $('lanInfo').textContent = data.lan_url ? `모바일 접속: ${data.lan_url}` : '';

      const hasKey = Boolean($('apiKeyInput').value.trim() || (data.settings && data.settings.api_key_set));
      const badge = $('apiBadge');
      badge.textContent = hasKey ? 'KCSC API 준비됨' : 'API 키 미입력';
      badge.className = hasKey ? 'badge' : 'badge warn';

      const total = data.progress_total || 0;
      const done = data.progress_done || 0;
      const pct = total > 0 ? Math.round((done / total) * 100) : 0;
      $('progressBar').style.width = `${pct}%`;
      $('progressText').textContent = `${pct}% (${done} / ${total})`;

      const logBox = $('logBox');
      logBox.textContent = (data.logs || []).join('\n') || '대기 중...';
      logBox.scrollTop = logBox.scrollHeight;

      $('runBtn').disabled = Boolean(data.running);
      $('runBtn').textContent = data.running ? '수집 및 워드 생성 중...' : '시방서 자동 생성 실행';

      const dlBtn = $('downloadBtn');
      if (data.download_ready && !data.running) {
        dlBtn.style.display = 'inline-flex';
        dlBtn.textContent = `${data.download_filename || '완성_시방서.docx'} 다운로드`;
      } else {
        dlBtn.style.display = 'none';
      }
      const dlXlsxBtn = $('downloadExcelBtn');
      if (data.excel_download_ready && !data.running) {
        dlXlsxBtn.style.display = 'inline-flex';
      } else {
        dlXlsxBtn.style.display = 'none';
      }

      const errBanner = $('errorBanner');
      if (data.last_error) {
        errBanner.style.display = 'block';
        errBanner.textContent = data.last_error;
      } else {
        errBanner.style.display = 'none';
      }

      if (data.running && !pollTimer) {
        pollTimer = setInterval(pollStatus, 700);
      } else if (!data.running && pollTimer) {
        clearInterval(pollTimer);
        pollTimer = null;
      }
    }

    async function pollStatus() {
      const res = await fetch('/api/state');
      const data = await res.json();
      applyServerState(data, false);
    }

    function renderCodes() {
      const q = $('codeSearch').value.trim().toLowerCase();
      const tbody = $('codeTableBody');
      tbody.innerHTML = '';
      let activeCount = 0;
      appState.codes.forEach((c, idx) => {
        if (c.include) activeCount++;
        const hay = `${c.display_code} ${c.title} ${c.category}`.toLowerCase();
        if (q && !hay.includes(q)) return;
        const tr = document.createElement('tr');
        tr.innerHTML = `
          <td><input type="checkbox" class="toggle" ${c.include ? 'checked' : ''} data-idx="${idx}" /></td>
          <td style="font-family:monospace; color:var(--cyan);">${c.display_code}</td>
          <td>${escapeHtml(c.title)}</td>
          <td class="sub">${escapeHtml(c.category)}</td>
        `;
        tr.querySelector('input').addEventListener('change', (e) => {
          appState.codes[idx].include = e.target.checked;
          updateCodeCount();
        });
        tbody.appendChild(tr);
      });
      $('codeCount').textContent = activeCount;
    }

    function updateCodeCount() {
      $('codeCount').textContent = appState.codes.filter(c => c.include).length;
    }

    function renderRules() {
      const tbody = $('ruleTableBody');
      tbody.innerHTML = '';
      appState.rules.forEach((r, idx) => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
          <td><input type="text" value="${escapeHtml(r.find_text)}" data-k="find_text" /></td>
          <td><input type="text" value="${escapeHtml(r.replace_text)}" placeholder="비워두면 삭제" data-k="replace_text" /></td>
          <td><input type="checkbox" class="toggle" ${r.is_regex ? 'checked' : ''} /></td>
          <td><button class="btn btn-ghost" style="padding:4px 8px;">삭제</button></td>
        `;
        const inputs = tr.querySelectorAll('input[type="text"]');
        inputs[0].addEventListener('input', (e) => appState.rules[idx].find_text = e.target.value);
        inputs[1].addEventListener('input', (e) => appState.rules[idx].replace_text = e.target.value);
        tr.querySelector('input[type="checkbox"]').addEventListener('change', (e) => appState.rules[idx].is_regex = e.target.checked);
        tr.querySelector('button').addEventListener('click', () => {
          appState.rules.splice(idx, 1);
          renderRules();
        });
        tbody.appendChild(tr);
      });
      $('ruleCount').textContent = appState.rules.length;
    }

    function escapeHtml(s) {
      return String(s || '').replace(/[&<>"']/g, (m) => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]));
    }

    // 엑셀 업로드
    const dz = $('dropzone');
    const fi = $('fileInput');
    dz.addEventListener('click', () => fi.click());
    dz.addEventListener('dragover', (e) => { e.preventDefault(); dz.classList.add('drag'); });
    dz.addEventListener('dragleave', () => dz.classList.remove('drag'));
    dz.addEventListener('drop', (e) => {
      e.preventDefault(); dz.classList.remove('drag');
      if (e.dataTransfer.files[0]) uploadExcel(e.dataTransfer.files[0]);
    });
    fi.addEventListener('change', () => { if (fi.files[0]) uploadExcel(fi.files[0]); });

    async function uploadExcel(file) {
      const reader = new FileReader();
      reader.onload = async () => {
        const b64 = reader.result.split(',')[1];
        const res = await fetch('/api/upload-excel', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ filename: file.name, content_base64: b64 })
        });
        const data = await res.json();
        if (!res.ok) {
          alert(data.error || '엑셀 업로드 실패');
          return;
        }
        applyServerState(data, true);
      };
      reader.readAsDataURL(file);
    }

    $('codeSearch').addEventListener('input', renderCodes);
    $('selectAllBtn').addEventListener('click', () => { appState.codes.forEach(c => c.include = true); renderCodes(); });
    $('clearAllBtn').addEventListener('click', () => { appState.codes.forEach(c => c.include = false); renderCodes(); });
    $('addRuleBtn').addEventListener('click', () => {
      appState.rules.push({ find_text: '', replace_text: '', is_regex: false, case_sensitive: true });
      renderRules();
    });
    $('apiKeyInput').addEventListener('input', () => applyServerState({ ...appState, logs: $('logBox').textContent.split('\n') }, false));

    $('runBtn').addEventListener('click', async () => {
      const payload = {
        codes: appState.codes,
        rules: appState.rules,
        allow_partial: $('allowPartialToggle').checked,
        settings: {
          api_key: $('apiKeyInput').value.trim(),
          project_title: $('projectTitleInput').value.trim(),
          output_filename: $('outputFilenameInput').value.trim(),
          page_break_per_code: $('pageBreakToggle').checked,
        }
      };
      const res = await fetch('/api/generate', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
      const data = await res.json();
      if (!res.ok) {
        alert(data.error || '실행 실패');
        return;
      }
      applyServerState(data, false);
    });

    fetchState();
  </script>
</body>
</html>
"""


class RequestHandler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        pass  # 콘솔 로그 깔끔하게 유지

    def _send_json(self, payload: Any, status: int = 200):
        raw = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(raw)))
        self.end_headers()
        self.wfile.write(raw)

    def _read_json_body(self) -> Dict[str, Any]:
        length = int(self.headers.get("Content-Length", "0"))
        if length <= 0:
            return {}
        return json.loads(self.rfile.read(length).decode("utf-8"))

    def do_GET(self):
        path = urllib.parse.urlparse(self.path).path
        if path == "/":
            raw = HTML_PAGE.encode("utf-8")
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(raw)))
            self.end_headers()
            self.wfile.write(raw)
        elif path == "/manifest.webmanifest":
            raw = MANIFEST_JSON.encode("utf-8")
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "application/manifest+json; charset=utf-8")
            self.send_header("Content-Length", str(len(raw)))
            self.end_headers()
            self.wfile.write(raw)
        elif path == "/icon.svg":
            raw = ICON_SVG.encode("utf-8")
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "image/svg+xml; charset=utf-8")
            self.send_header("Content-Length", str(len(raw)))
            self.end_headers()
            self.wfile.write(raw)
        elif path == "/api/state":
            self._send_json(STATE.snapshot())
        elif path == "/api/download-excel":
            with STATE.lock:
                xlsx_out = STATE.last_excel_path
            if not xlsx_out or not xlsx_out.exists():
                self._send_json({"error": "저장된 엑셀 파일이 없습니다."}, status=404)
                return
            raw = xlsx_out.read_bytes()
            encoded_name = urllib.parse.quote(xlsx_out.name)
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Disposition", f"attachment; filename*=UTF-8''{encoded_name}")
            self.send_header("Content-Length", str(len(raw)))
            self.end_headers()
            self.wfile.write(raw)
        elif path == "/api/download":
            with STATE.lock:
                out_path = STATE.last_output_path
            if not out_path or not out_path.exists():
                self._send_json({"error": "생성된 워드 파일이 없습니다."}, status=404)
                return
            raw = out_path.read_bytes()
            encoded_name = urllib.parse.quote(out_path.name)
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            self.send_header("Content-Disposition", f"attachment; filename*=UTF-8''{encoded_name}")
            self.send_header("Content-Length", str(len(raw)))
            self.end_headers()
            self.wfile.write(raw)
        else:
            self._send_json({"error": "Not found"}, status=404)

    def do_POST(self):
        path = urllib.parse.urlparse(self.path).path
        if path == "/api/upload-excel":
            try:
                body = self._read_json_body()
                filename = pathlib.Path(body.get("filename") or "uploaded.xlsx").name
                b64 = body.get("content_base64") or ""
                if not b64:
                    raise ValueError("파일 내용이 비어 있습니다.")
                tmp_path = OUTPUT_DIR / f"_uploaded_{filename}"
                tmp_path.write_bytes(base64.b64decode(b64))
                STATE.load_excel_path(tmp_path, display_name=filename)
                self._send_json(STATE.snapshot())
            except Exception as exc:
                self._send_json({"error": str(exc)}, status=400)
        elif path == "/api/generate":
            try:
                body = self._read_json_body()
                with STATE.lock:
                    if STATE.running:
                        self._send_json({"error": "이미 시방서 생성이 진행 중입니다."}, status=409)
                        return
                t = threading.Thread(target=run_generation_job, args=(body,), daemon=True)
                t.start()
                time.sleep(0.05)
                self._send_json(STATE.snapshot())
            except Exception as exc:
                self._send_json({"error": str(exc)}, status=400)
        else:
            self._send_json({"error": "Not found"}, status=404)


def main():
    global SERVER_PORT
    parser = argparse.ArgumentParser(description="KCSC 시방서 자동화 프로페셔널 다크모드 웹/PWA 앱")
    parser.add_argument("--host", default="0.0.0.0", help="바인딩 호스트 (기본값: 0.0.0.0)")
    parser.add_argument("--port", "-p", type=int, default=8501, help="포트 번호 (기본값: 8501)")
    parser.add_argument("--no-browser", action="store_true", help="실행 시 브라우저 자동 열기 비활성화")
    args = parser.parse_args()
    SERVER_PORT = args.port

    default_xlsx = BASE_DIR / "시방서_설정.xlsx"
    if default_xlsx.exists():
        try:
            STATE.load_excel_path(default_xlsx)
        except Exception as exc:
            STATE.add_log(f"기본 시방서_설정.xlsx 자동 로드 실패: {exc}")

    server = ThreadingHTTPServer((args.host, args.port), RequestHandler)
    local_url = f"http://127.0.0.1:{args.port}"
    lan_url = f"http://{_get_lan_ip()}:{args.port}"
    print("=" * 60)
    print("KCSC 시방서 자동화 (Pro Dark Mode) 실행 중")
    print(f"  - PC 접속/앱 설치: {local_url}")
    print(f"  - 스마트폰(같은 Wi-Fi): {lan_url}")
    print("종료하려면 Ctrl+C를 누르세요.")
    print("=" * 60)

    if not args.no_browser:
        threading.Timer(0.6, lambda: webbrowser.open(local_url)).start()

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n서버를 종료합니다.")
        server.server_close()


if __name__ == "__main__":
    main()