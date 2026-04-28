"""マニュアルチェックリスト生成用プロンプト (シート単位 AI 呼び出し)。"""
from __future__ import annotations

import json

from ...parser.models import Sheet


CATEGORIES = ["記入漏れ", "整合性"]


SYSTEM = """あなたは臨床試験の Data Manager (DM) です。SDV 前に DM が目視確認すべきチェックポイントを抽出してください。

# 出力カテゴリ
- 記入漏れ: EDC で必須化されていないが、条件付きで入力されているべき項目の入力漏れ確認
- 整合性: 単一フォーム内の項目間整合性 (日付ペアの前後・選択肢と他項目の連動・SAE 報告の整合 など)

# 重要なルール
- target_field は `ラベル(field名)` の形式で 1 件のみ指定してください。
- 出力は必ず指定された JSON のみを出力してください。Markdown の修飾 (```json など) や前後の挨拶は一切不要です。
- severity は "high", "medium", "low" のいずれかを指定してください。

# 出力例
{
  "check_points": [
    {
      "category": "整合性",
      "target_field": "同意取得日(field1)",
      "check_point": "同意取得日が生年月日より後か確認する。",
      "rationale": "年齢要件の確認のため。",
      "severity": "high"
    }
  ]
}
"""


def _candidate_items(sheet: Sheet) -> list[dict]:
    items: list[dict] = []
    for fi in sheet.field_items:
        if fi.type == "FieldItem::Note":
            continue
        if not fi.field_type:
            continue
        if fi.field_type in ("text", "drug", "meddra"):
            continue
        v = fi.validators
        items.append(
            {
                "field": fi.name,
                "label": fi.label,
                "type": fi.field_type,
                "required": v.presence is not None,
                "has_default": bool(fi.default_value),
                "has_numericality": v.numericality is not None,
            }
        )
    return items


def build_user_prompt(sheet: Sheet) -> str:
    items = _candidate_items(sheet)
    payload = {"sheet_name": sheet.name, "candidate_items": items}
    return (
        "次の CRF フォームについて、DM が SDV 前に目視確認すべきチェックポイントを抽出してください。\n"
        "観点は『記入漏れ』『整合性』の 2 カテゴリのみとし、該当がない場合は空の配列を返してください。\n\n"
        f"{json.dumps(payload, ensure_ascii=False, indent=2)}"
    )


SCHEMA = {
    "type": "object",
    "properties": {
        "check_points": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "category": {"type": "string"},
                    "target_field": {"type": "string"},
                    "check_point": {"type": "string"},
                    "rationale": {"type": "string"},
                    "severity": {"type": "string"},
                },
                "required": ["category", "target_field", "check_point", "rationale", "severity"],
            },
        }
    },
    "required": ["check_points"],
}