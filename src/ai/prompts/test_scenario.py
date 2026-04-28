"""テストシナリオ生成用プロンプト。"""
from __future__ import annotations

import json

from ...parser.models import Sheet, Study


SYSTEM = """あなたは臨床試験の EDC システムを検証するシニア QA エンジニアです。
入力された CRF フォームの設定情報をもとに、境界値分析と同値分割に基づくテストシナリオを生成してください。
出力は必ず指定された JSON スキーマに従い、日本語で記述してください。

# ルール
1. input_value には CRF 画面上でユーザーが選ぶ/入力する表示文字列 (label/name) を記述すること。
   - コードリスト項目では code (例: "MALE") ではなく **表示名** (例: "男性") を使用する。
   - 日付は "YYYY-MM-DD" 形式、数値はそのままの値を文字列として記述する。
2. コードリスト (code_list_values が与えられている項目) については、
   **「コードリスト範囲外の値」を異常系として生成しないこと**。
   EDC の UI 制約で選択肢以外は入力できないため、テスト不要。
   - ただし、必須未入力、formula 違反などの異常系は生成してよい。
3. それ以外の項目 (text/数値/日付/自由記述) については、境界値分析・同値分割により
   正常系 (kind=normal) と異常系 (kind=abnormal) をそれぞれ最低 1 件ずつ生成する。
4. expected_result は具体的に (例: "正常に保存される", "エラーメッセージ「同意取得時年齢をご確認下さい」が表示される" など)。
"""


def _codelist_values(study: Study | None, option_name: str | None) -> list[dict] | None:
    if study is None or not option_name:
        return None
    cl = study.codelist_by_name(option_name)
    if cl is None:
        return None
    return [{"code": v.code, "name": v.name} for v in cl.values if v.is_usable]


def build_user_prompt(sheet: Sheet, study: Study | None = None) -> str:
    """フォーム 1 シート分の入力情報を文字列化。"""
    items = []
    for fi in sheet.field_items:
        if fi.type in ("FieldItem::Note", "FieldItem::Heading", "FieldItem::Allocation"):
            continue
        if fi.is_invisible:
            continue
        v = fi.validators
        items.append(
            {
                "field": fi.name,
                "label": fi.label,
                "type": fi.field_type or fi.type.replace("FieldItem::", ""),
                "required": v.presence is not None,
                "code_list": fi.option_name,
                "code_list_values": _codelist_values(study, fi.option_name),
                "numericality": v.numericality.model_dump(exclude_none=True) if v.numericality else None,
                "date": v.date.model_dump(exclude_none=True) if v.date else None,
                "formula": v.formula.model_dump(exclude_none=True) if v.formula else None,
            }
        )
    payload = {"sheet_name": sheet.name, "items": items}
    return (
        "次の CRF フォーム設定について、テストシナリオを生成してください。\n\n"
        f"```json\n{json.dumps(payload, ensure_ascii=False, indent=2)}\n```\n\n"
        "各 item につき、最低 1 件の正常系 (kind=normal) を含めてください。"
        "ただしコードリスト項目の異常系は不要 (UI で制約済み) です。"
    )


SCHEMA = {
    "type": "object",
    "properties": {
        "scenarios": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "field": {"type": "string"},
                    "label": {"type": "string"},
                    "kind": {"type": "string", "enum": ["normal", "abnormal"]},
                    "input_value": {"type": "string"},
                    "expected_result": {"type": "string"},
                    "rationale": {"type": "string"},
                },
                "required": ["field", "kind", "input_value", "expected_result"],
            },
        }
    },
    "required": ["scenarios"],
}
