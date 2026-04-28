"""マニュアルチェックリスト (Excel) 生成 — Cohere 利用。"""
from __future__ import annotations

import re
from datetime import date
from pathlib import Path
from typing import Callable

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from ..ai.cohere_client import CohereJSONClient
from ..ai.prompts import manual_check as P
from ..parser.models import Study

BASE_FONT = "Meiryo UI"
F_BASE = Font(name=BASE_FONT, size=9)
F_HEADER = Font(name=BASE_FONT, size=10, bold=True, color="FFFFFF")
F_TITLE = Font(name=BASE_FONT, size=16, bold=True)
FILL_HEADER = PatternFill("solid", fgColor="305496")
FILL_HIGH = PatternFill("solid", fgColor="F8CBAD")
FILL_MID = PatternFill("solid", fgColor="FFE699")
FILL_LOW = PatternFill("solid", fgColor="E2EFDA")
THIN = Side(style="thin", color="BFBFBF")
BORDER_ALL = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(vertical="top", wrap_text=True)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)

HEADERS = ["No.", "Sheet", "Category", "Target Fields", "Check Point", "Rationale", "Severity"]
WIDTHS = [6, 24, 16, 24, 50, 36, 10]

CATEGORY_ORDER = {"記入漏れ": 0, "単位・桁数": 1, "整合性": 2}


def _estimate_row_height(values: list, widths: list[int]) -> float:
    max_lines = 1
    for v, w in zip(values, widths):
        text = str(v) if v is not None else ""
        if not text:
            continue
        for line in text.replace("\r", "").split("\n"):
            est_chars = max(1, w - 1)
            wrap_lines = max(1, -(-int(len(line) * 1.5) // est_chars))
            max_lines = max(max_lines, wrap_lines)
    return max(15.0, min(150.0, max_lines * 14.0))


def _severity_fill(sev: str) -> PatternFill | None:
    return {"high": FILL_HIGH, "medium": FILL_MID, "low": FILL_LOW}.get(sev)


def generate_check_points(
    study: Study,
    selected_sheet_names: list[str],
    client: CohereJSONClient | None = None,
    on_progress: Callable[[int, int, str], None] | None = None,
) -> list[dict]:
    """シート単位で AI を呼び出し、チェックポイントを生成。重複と必須項目を自動除外。"""
    target_sheets = [s for s in study.sheets if s.name in set(selected_sheet_names)]
    out: list[dict] = []
    total = len(target_sheets)
    client = client or CohereJSONClient()

    # 全体での重複排除用
    seen_keys = set()

    ai_skip = False
    consecutive_errors = 0
    
    for i, sheet in enumerate(target_sheets, start=1):
        if on_progress:
            on_progress(i, total, sheet.name)
        
        # このシートの全フィールド情報を辞書化
        field_map = {fi.name: fi for fi in sheet.field_items}
        valid_field_names = set(field_map.keys())

        if not ai_skip:
            try:
                data = client.chat_json(P.SYSTEM, P.build_user_prompt(sheet), P.SCHEMA)
                consecutive_errors = 0
                
                for cp in data.get("check_points", []):
                    category = cp.get("category", "")
                    if category not in P.CATEGORIES:
                        continue
                    
                    tgt_str = cp.get("target_field", "")
                    if not tgt_str:
                        continue

                    # 1. 必須チェック済み項目の除外 (Python側で二重ガード)
                    if category == "記入漏れ":
                        match = re.search(r"field\d+", tgt_str)
                        if match:
                            f_name = match.group()
                            fi = field_map.get(f_name)
                            # 必須設定 (presence) がある場合はマニュアルチェック不要
                            if fi and (fi.validators.presence is not None or fi.default_value):
                                continue

                    # 2. 重複排除 (同一シート、同一項目、同一内容)
                    check_text = cp.get("check_point", "")
                    unique_key = (sheet.name, tgt_str, check_text)
                    if unique_key in seen_keys:
                        continue

                    # 存在しないフィールドへの言及チェック
                    referenced = re.findall(r"field\d+", tgt_str)
                    if not referenced or any(f not in valid_field_names for f in referenced):
                        continue
                        
                    out.append({
                        "sheet": sheet.name,
                        "category": category,
                        "target_fields": [tgt_str],
                        "check_point": check_text,
                        "rationale": cp.get("rationale", ""),
                        "severity": cp.get("severity", "medium"),
                    })
                    seen_keys.add(unique_key)

            except Exception as e:
                consecutive_errors += 1
                out.append({
                    "sheet": sheet.name, "category": "(error)", "target_fields": [],
                    "check_point": f"AI生成エラー: {e}", "rationale": "", "severity": "low",
                })
                if consecutive_errors >= 2:
                    ai_skip = True

        # 決定論的な「単位・桁数」チェックの追加
        for unit_check in _unit_digit_checks_for_sheet(sheet):
            tgt = unit_check["target_fields"][0]
            ukey = (sheet.name, tgt, unit_check["check_point"])
            if ukey not in seen_keys:
                out.append(unit_check)
                seen_keys.add(ukey)

    sheet_order = {s.name: i for i, s in enumerate(study.sheets)}
    out.sort(key=lambda cp: (sheet_order.get(cp.get("sheet", ""), 999), CATEGORY_ORDER.get(cp.get("category", ""), 99)))

    return _expand_one_target_per_row(out)


def _unit_digit_checks_for_sheet(sheet) -> list[dict]:
    out: list[dict] = []
    for fi in sheet.field_items:
        if fi.type == "FieldItem::Note" or not fi.field_type:
            continue
        if not fi.validators.numericality:
            continue
        out.append({
            "sheet": sheet.name,
            "category": "単位・桁数",
            "target_fields": [f"{fi.label}({fi.name})"],
            "check_point": f"{fi.label} の単位・桁数の妥当性を確認する (例: mg/g 取り違え、桁ずれ)。",
            "rationale": "数値範囲内でも単位誤りはシステムで検出できないため。",
            "severity": "medium",
        })
    return out


def _expand_one_target_per_row(rows: list[dict]) -> list[dict]:
    expanded: list[dict] = []
    for cp in rows:
        targets = cp.get("target_fields") or []
        if not isinstance(targets, list) or len(targets) <= 1:
            expanded.append(cp)
            continue
        for t in targets:
            new_cp = cp.copy()
            new_cp["target_fields"] = [t]
            expanded.append(new_cp)
    return expanded


def build_manual_check_workbook(study: Study, points: list[dict]) -> Workbook:
    wb = Workbook()
    wb.remove(wb.active)
    cv = wb.create_sheet("表紙")
    cv.sheet_view.showGridLines = False
    cv.column_dimensions["B"].width = 22
    cv.column_dimensions["C"].width = 60
    cv.merge_cells("B3:C3")
    t = cv["B3"]
    t.value = "マニュアルチェックリスト"
    t.font = Font(name=BASE_FONT, size=24, bold=True, color="1F3864")
    t.alignment = CENTER
    cv.merge_cells("B5:C5")
    pn = cv["B5"]
    pn.value = study.proper_name
    pn.font = Font(name=BASE_FONT, size=12, color="404040")
    pn.alignment = CENTER
    metas = [("試験 ID", study.name), ("チェック件数", f"{len(points):,}"), ("発行日", date.today().isoformat())]
    for i, (k, v) in enumerate(metas):
        r = 9 + i
        cv.cell(row=r, column=2, value=k).font = Font(name=BASE_FONT, size=10, bold=True)
        cv.cell(row=r, column=3, value=v).font = F_BASE

    ws = wb.create_sheet("マニュアルチェック一覧")
    ws.sheet_view.showGridLines = False
    ws["A1"] = "マニュアルチェック一覧"
    ws["A1"].font = F_TITLE
    header_row = 3
    for i, h in enumerate(HEADERS, start=1):
        c = ws.cell(row=header_row, column=i, value=h)
        c.font = F_HEADER
        c.fill = FILL_HEADER
        c.alignment = CENTER
        c.border = BORDER_ALL
        ws.column_dimensions[get_column_letter(i)].width = WIDTHS[i - 1]

    for idx, cp in enumerate(points, start=1):
        r = header_row + idx
        tgts = ", ".join(cp.get("target_fields", []))
        values = [idx, cp.get("sheet", ""), cp.get("category", ""), tgts, cp.get("check_point", ""), cp.get("rationale", ""), cp.get("severity", "")]
        for j, v in enumerate(values, start=1):
            c = ws.cell(row=r, column=j, value=v)
            c.font = F_BASE
            c.border = BORDER_ALL
            c.alignment = WRAP
        fill = _severity_fill(cp.get("severity", ""))
        if fill:
            ws.cell(row=r, column=7).fill = fill
        ws.row_dimensions[r].height = _estimate_row_height(values, WIDTHS)

    if points:
        ws.auto_filter.ref = f"A{header_row}:{get_column_letter(len(HEADERS))}{header_row + len(points)}"
    ws.freeze_panes = ws.cell(row=header_row + 1, column=1)
    return wb

def write_manual_check_excel(study: Study, selected_sheet_names: list[str], output_path: str | Path, client: CohereJSONClient | None = None, on_progress: Callable[[int, int, str], None] | None = None) -> Path:
    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    points = generate_check_points(study, selected_sheet_names, client, on_progress)
    wb = build_manual_check_workbook(study, points)
    wb.save(out)
    return out