"""マニュアルチェックリスト生成用プロンプト。"""
from __future__ import annotations

import json

from ...parser.models import Sheet


SYSTEM = """あなたは臨床試験の Data Manager (DM) です。
EDC システム上で機械的な制約 (validators) ではカバーできない、目視確認すべきポイントを抽出してください。
特に注目すべきポイント:
- 自由記述項目 (text) における表記ゆれ・記入漏れ
- 併用薬と有害事象、AE と原疾患の整合性 など、フォーム横断の妥当性
- SOC/PT (MedDRA) のコーディング妥当性
- 単位や桁数、日付の整合性

出力は必ず指定された JSON スキーマに従い、日本語で記述してください。

# 重要なルール
- target_fields には **必ず 1 つ以上の具体的な対象フィールドを `ラベル(field名)` の形式 (例: `投与量(field3)`) で指定**してください。
- 対象フィールドが特定できないチェックポイントは出力しないでください (汎用的な注意喚起は不要)。
- Note (注釈) 項目自体はチェック対象外とし、target_fields に含めないでください。
"""


def build_user_prompt(sheet: Sheet) -> str:
    items = []
    for fi in sheet.field_items:
        # Note 項目はチェック対象外
        if fi.type == "FieldItem::Note":
            continue
        # 制約がない/弱い項目を中心に拾う
        v = fi.validators
        is_freetext = (fi.field_type == "text") and not v.numericality and not v.formula
        if not (is_freetext or fi.field_type in ("drug", "meddra", "sae_report")):
            continue
        items.append(
            {
                "field": fi.name,
                "label": fi.label,
                "type": fi.field_type or fi.type.replace("FieldItem::", ""),
                "description": fi.description or None,
            }
        )
    payload = {"sheet_name": sheet.name, "candidate_items": items}
    return (
        "次の CRF フォームについて、DM 担当者が目視確認すべきチェックポイントを抽出してください。\n"
        "機械的な validator では検出できない問題を中心に、最低 3 件挙げてください。\n\n"
        f"```json\n{json.dumps(payload, ensure_ascii=False, indent=2)}\n```"
    )


SCHEMA = {
    "type": "object",
    "properties": {
        "check_points": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "category": {"type": "string", "description": "整合性 / 表記 / コーディング 等"},
                    "target_fields": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "対象フィールドを `ラベル(field名)` の形式 (例: `投与量(field3)`) で 1 件以上指定すること。",
                    },
                    "check_point": {"type": "string"},
                    "rationale": {"type": "string"},
                    "severity": {"type": "string", "enum": ["high", "medium", "low"]},
                },
                "required": ["category", "target_fields", "check_point", "severity"],
            },
        }
    },
    "required": ["check_points"],
}
