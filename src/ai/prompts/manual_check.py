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
- **【最重要】必須化されている項目の除外**: candidate_items の中で **`required: true` または `has_default: true` となっている項目に対して「記入漏れ」のチェックを作成しないでください。** これらはシステムで制御されているため、目視確認リストに含める必要はありません。
- target_field は `ラベル(field名)` の形式で 1 件のみ指定してください。
- 出力は必ず指定された JSON のみを出力してください。Markdown の修飾や挨拶は一切不要です。
- 各シート 1〜5 件程度に絞り、該当がない場合は空の配列 `[]` を返してください。

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
                "required": v.presence is not None, # AIがこれを見て判断します
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
        "『required: true』の項目に対する記入漏れチェックは絶対に含めないでください。\n\n"
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