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
    """シート単位で AI を呼び『記入漏れ / 整合性』を抽出し、決定論的に『単位・桁数』を追加。"""
    target_sheets = [s for s in study.sheets if s.name in set(selected_sheet_names)]
    out: list[dict] = []
    total = len(target_sheets)
    client = client or CohereJSONClient()

    ai_skip = False
    consecutive_errors = 0
    for i, sheet in enumerate(target_sheets, start=1):
        if on_progress:
            on_progress(i, total, sheet.name)
        valid_fields = {fi.name for fi in sheet.field_items if fi.type != "FieldItem::Note"}

        # AI: 記入漏れ + 整合性 (連続エラー時は以降スキップして決定論パートだけ実行)
        if ai_skip:
            out.extend(_unit_digit_checks_for_sheet(sheet))
            continue
        try:
            data = client.chat_json(P.SYSTEM, P.build_user_prompt(sheet), P.SCHEMA)
            consecutive_errors = 0
            for cp in data.get("check_points", []):
                if cp.get("category") not in P.CATEGORIES:
                    continue
                tgt = cp.get("target_fields")
                if not tgt or (isinstance(tgt, list) and not [t for t in tgt if t]):
                    continue
                tgt_str = " ".join(tgt) if isinstance(tgt, list) else str(tgt)
                referenced = re.findall(r"field\d+", tgt_str)
                if not referenced or any(f not in valid_fields for f in referenced):
                    continue
                out.append(
                    {
                        "sheet": sheet.name,
                        "category": cp.get("category", ""),
                        "target_fields": tgt,
                        "check_point": cp.get("check_point", ""),
                        "rationale": cp.get("rationale", ""),
                        "severity": cp.get("severity", "medium"),
                    }
                )
        except Exception as e:
            consecutive_errors += 1
            out.append(
                {
                    "sheet": sheet.name,
                    "category": "(error)",
                    "target_fields": [],
                    "check_point": f"AI生成エラー: {e}",
                    "rationale": "",
                    "severity": "low",
                }
            )
            # 連続 2 件以上 AI が失敗した場合は以降の AI 呼び出しを諦める
            if consecutive_errors >= 2:
                ai_skip = True

        # 決定論: 単位・桁数
        out.extend(_unit_digit_checks_for_sheet(sheet))

    # 並び順: シート順 → カテゴリ順
    sheet_order = {s.name: i for i, s in enumerate(study.sheets)}
    out.sort(
        key=lambda cp: (
            sheet_order.get(cp.get("sheet", ""), 9999),
            CATEGORY_ORDER.get(cp.get("category", ""), 99),
        )
    )

    # 1 ターゲットフィールド = 1 行 に展開
    return _expand_one_target_per_row(out)


def _unit_digit_checks_for_sheet(sheet) -> list[dict]:
    """numericality 制約のある数値項目に単位・桁数チェックを生成。"""
    out: list[dict] = []
    for fi in sheet.field_items:
        if fi.type == "FieldItem::Note" or not fi.field_type:
            continue
        if fi.validators.numericality is None:
            continue
        out.append(
            {
                "sheet": sheet.name,
                "category": "単位・桁数",
                "target_fields": [f"{fi.label}({fi.name})"],
                "check_point": f"{fi.label} の単位・桁数の妥当性を確認する (例: mg/g 取り違え、桁ずれ)。",
                "rationale": "numericality 範囲内でも単位・桁数の入力誤りはバリデーションで検出できないため。",
                "severity": "medium",
            }
        )
    return out


def _expand_one_target_per_row(rows: list[dict]) -> list[dict]:
    expanded: list[dict] = []
    for cp in rows:
        targets = cp.get("target_fields") or []
        if not isinstance(targets, list) or len(targets) <= 1:
            expanded.append(cp)
            continue
        for t in targets:
            expanded.append({**cp, "target_fields": [t]})
    return expanded


def build_manual_check_workbook(study: Study, points: list[dict]) -> Workbook:
    wb = Workbook()
    wb.remove(wb.active)

    cv = wb.create_sheet("表紙")
    cv.sheet_view.showGridLines = False
    cv.column_dimensions["A"].width = 2
    cv.column_dimensions["B"].width = 22
    cv.column_dimensions["C"].width = 60
    cv.merge_cells("B3:C3")
    t = cv["B3"]
    t.value = "マニュアルチェックリスト"
    t.font = Font(name=BASE_FONT, size=24, bold=True, color="1F3864")
    t.alignment = CENTER
    cv.row_dimensions[3].height = 50
    cv.row_dimensions[5].height = 60
    cv.merge_cells("B5:C5")
    pn = cv["B5"]
    pn.value = study.proper_name
    pn.font = Font(name=BASE_FONT, size=12, color="404040")
    pn.alignment = CENTER
    metas = [("試験 ID", study.name), ("チェック件数", f"{len(points):,}"), ("発行日", date.today().isoformat())]
    for i, (k, v) in enumerate(metas):
        r = 9 + i
        cv.row_dimensions[r].height = 22
        kc = cv.cell(row=r, column=2, value=k)
        kc.font = Font(name=BASE_FONT, size=10, bold=True)
        kc.fill = PatternFill("solid", fgColor="D9E1F2")
        kc.border = BORDER_ALL
        kc.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        vc = cv.cell(row=r, column=3, value=v)
        vc.font = F_BASE
        vc.border = BORDER_ALL
        vc.alignment = Alignment(horizontal="left", vertical="center", indent=1)

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
        targets = cp.get("target_fields") or []
        if isinstance(targets, list):
            targets = ", ".join(targets)
        values = [
            idx,
            cp.get("sheet", ""),
            cp.get("category", ""),
            targets,
            cp.get("check_point", ""),
            cp.get("rationale", ""),
            cp.get("severity", ""),
        ]
        for j, v in enumerate(values, start=1):
            c = ws.cell(row=r, column=j, value=v)
            c.font = F_BASE
            c.border = BORDER_ALL
            c.alignment = WRAP
        fill = _severity_fill(cp.get("severity", ""))
        if fill is not None:
            ws.cell(row=r, column=7).fill = fill
        ws.row_dimensions[r].height = _estimate_row_height(values, WIDTHS)

    if points:
        last_col = get_column_letter(len(HEADERS))
        ws.auto_filter.ref = f"A{header_row}:{last_col}{header_row + len(points)}"
    ws.freeze_panes = ws.cell(row=header_row + 1, column=1)
    return wb


def write_manual_check_excel(
    study: Study,
    selected_sheet_names: list[str],
    output_path: str | Path,
    client: CohereJSONClient | None = None,
    on_progress: Callable[[int, int, str], None] | None = None,
) -> Path:
    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    points = generate_check_points(study, selected_sheet_names, client, on_progress)
    wb = build_manual_check_workbook(study, points)
    wb.save(out)
    return out
