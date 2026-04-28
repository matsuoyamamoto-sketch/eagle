"""テストシナリオ (Excel) 生成 — Cohere 利用。"""
from __future__ import annotations

from datetime import date
from pathlib import Path
from typing import Callable

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from ..ai.cohere_client import CohereJSONClient
from ..ai.prompts import test_scenario as P
from ..parser.models import Sheet, Study

BASE_FONT = "Meiryo UI"
F_BASE = Font(name=BASE_FONT, size=9)
F_HEADER = Font(name=BASE_FONT, size=10, bold=True, color="FFFFFF")
F_TITLE = Font(name=BASE_FONT, size=16, bold=True)
FILL_HEADER = PatternFill("solid", fgColor="305496")
FILL_NORMAL = PatternFill("solid", fgColor="E2EFDA")
FILL_ABNORMAL = PatternFill("solid", fgColor="FCE4D6")
THIN = Side(style="thin", color="BFBFBF")
BORDER_ALL = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(vertical="top", wrap_text=True)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)

HEADERS = ["No.", "Sheet", "Field", "Label", "Kind", "Input", "Expected", "Rationale"]
WIDTHS = [6, 24, 12, 24, 10, 28, 36, 40]


def _estimate_row_height(values: list, widths: list[int]) -> float:
    """各セルの折り返し行数を推定し、最大値から行高 (pt) を算出。"""
    max_lines = 1
    for v, w in zip(values, widths):
        text = str(v) if v is not None else ""
        if not text:
            continue
        # 改行を考慮
        for line in text.replace("\r", "").split("\n"):
            # 全角は概ね 2 char 幅相当として、列幅 (chars) を基準に行数推定
            est_chars = max(1, w - 1)
            # 日本語は1文字≒2幅。安全側で 1.5 倍に。
            wrap_lines = max(1, -(-int(len(line) * 1.5) // est_chars))
            max_lines = max(max_lines, wrap_lines)
    return max(15.0, min(150.0, max_lines * 14.0))


def _row_fill(kind: str) -> PatternFill | None:
    if kind == "normal":
        return FILL_NORMAL
    if kind == "abnormal":
        return FILL_ABNORMAL
    return None


def generate_scenarios(
    study: Study,
    selected_sheet_names: list[str],
    client: CohereJSONClient | None = None,
    on_progress: Callable[[int, int, str], None] | None = None,
) -> list[dict]:
    """指定フォームについて AI でシナリオ生成。"""
    client = client or CohereJSONClient()
    target = [s for s in study.sheets if s.name in selected_sheet_names]
    out: list[dict] = []
    total = len(target)
    for i, sheet in enumerate(target, start=1):
        if on_progress:
            on_progress(i, total, sheet.name)
        try:
            data = client.chat_json(P.SYSTEM, P.build_user_prompt(sheet, study), P.SCHEMA)
            for sc in data.get("scenarios", []):
                out.append({"sheet": sheet.name, **sc})
        except Exception as e:
            out.append(
                {
                    "sheet": sheet.name,
                    "field": "(error)",
                    "label": "",
                    "kind": "abnormal",
                    "input_value": "",
                    "expected_result": "",
                    "rationale": f"生成エラー: {e}",
                }
            )
    return out


def build_test_scenario_workbook(
    study: Study, scenarios: list[dict]
) -> Workbook:
    wb = Workbook()
    wb.remove(wb.active)

    # 表紙 (簡略)
    cv = wb.create_sheet("表紙")
    cv.sheet_view.showGridLines = False
    cv.column_dimensions["A"].width = 2
    cv.column_dimensions["B"].width = 22
    cv.column_dimensions["C"].width = 60
    cv.merge_cells("B3:C3")
    t = cv["B3"]
    t.value = "テストシナリオ"
    t.font = Font(name=BASE_FONT, size=24, bold=True, color="1F3864")
    t.alignment = CENTER
    cv.row_dimensions[3].height = 50
    cv.row_dimensions[5].height = 60
    cv.merge_cells("B5:C5")
    pn = cv["B5"]
    pn.value = study.proper_name
    pn.font = Font(name=BASE_FONT, size=12, color="404040")
    pn.alignment = CENTER
    metas = [("試験 ID", study.name), ("シナリオ件数", f"{len(scenarios):,}"), ("発行日", date.today().isoformat())]
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

    # 一覧シート
    ws = wb.create_sheet("テストシナリオ一覧")
    ws.sheet_view.showGridLines = False
    ws["A1"] = "テストシナリオ一覧"
    ws["A1"].font = F_TITLE

    header_row = 3
    for i, h in enumerate(HEADERS, start=1):
        c = ws.cell(row=header_row, column=i, value=h)
        c.font = F_HEADER
        c.fill = FILL_HEADER
        c.alignment = CENTER
        c.border = BORDER_ALL
        ws.column_dimensions[get_column_letter(i)].width = WIDTHS[i - 1]

    for idx, sc in enumerate(scenarios, start=1):
        r = header_row + idx
        values = [
            idx,
            sc.get("sheet", ""),
            sc.get("field", ""),
            sc.get("label", ""),
            sc.get("kind", ""),
            sc.get("input_value", ""),
            sc.get("expected_result", ""),
            sc.get("rationale", ""),
        ]
        for j, v in enumerate(values, start=1):
            c = ws.cell(row=r, column=j, value=v)
            c.font = F_BASE
            c.border = BORDER_ALL
            c.alignment = WRAP
        fill = _row_fill(sc.get("kind", ""))
        if fill is not None:
            ws.cell(row=r, column=5).fill = fill
        ws.row_dimensions[r].height = _estimate_row_height(values, WIDTHS)

    last_col = get_column_letter(len(HEADERS))
    ws.auto_filter.ref = f"A{header_row}:{last_col}{header_row + len(scenarios)}"
    ws.freeze_panes = ws.cell(row=header_row + 1, column=1)

    return wb


def write_test_scenario_excel(
    study: Study,
    selected_sheet_names: list[str],
    output_path: str | Path,
    client: CohereJSONClient | None = None,
    on_progress: Callable[[int, int, str], None] | None = None,
) -> Path:
    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    scenarios = generate_scenarios(study, selected_sheet_names, client, on_progress)
    wb = build_test_scenario_workbook(study, scenarios)
    wb.save(out)
    return out
