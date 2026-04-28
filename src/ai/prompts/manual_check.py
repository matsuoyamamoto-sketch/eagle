"""マニュアルチェックリスト生成用プロンプト。"""
from __future__ import annotations

import json

from ...parser.models import Sheet, Study


SYSTEM = """あなたは臨床試験の Data Manager (DM) です。SDV 前に DM が目視確認すべきチェックポイントを抽出してください。

# 出力カテゴリ (この 3 つのみ使用)
- 記入漏れ: EDC の必須設定では捕まらないが、条件付きで入力されているべき項目の入力漏れ確認
- 単位・桁数: 数値項目の単位・桁数の妥当性 (numericality 範囲内でも臨床的に疑わしい値の確認)
- 整合性: 単一フォーム内の項目間整合性 (日付ペアの前後・選択肢と他項目の連動・SAE 報告の整合 など)

# 重要なルール
- target_fields は **`ラベル(field名)` 形式** (例: `投与量(field3)`) で 1 件以上指定。
- target_fields に指定できるのは、**candidate_items に含まれる field/label のみ**です。存在しないフィールドを推測・創作しないでください。
- 上記 3 カテゴリ以外は出力しないでください。フォーム横断の整合性は別途決定論的に生成するため、出力に含めないでください。
- 自由記述項目 (text)・薬剤コーディング (drug)・MedDRA コーディング (meddra)・Note は対象外です (candidate_items に含まれません)。
- `has_default: true` の項目は初期値が入力済みのため、記入漏れチェックの対象に含めないでください。
- 出力は JSON スキーマ準拠、日本語で記述してください。
- 該当するチェックがない場合は check_points を空配列にしてください。
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
                "description": fi.description or None,
            }
        )
    return items


def build_user_prompt(sheet: Sheet, study: Study | None = None) -> str:
    items = _candidate_items(sheet)
    payload = {"sheet_name": sheet.name, "candidate_items": items}
    return (
        "次の CRF フォームについて、DM が SDV 前に目視確認すべきチェックポイントを抽出してください。\n"
        "観点は『記入漏れ / 単位・桁数 / 整合性』の 3 カテゴリのみです。\n\n"
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
                    "category": {
                        "type": "string",
                        "enum": ["記入漏れ", "単位・桁数", "整合性"],
                    },
                    "target_fields": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "対象フィールドを `ラベル(field名)` 形式 (例: `投与量(field3)`) で 1 件以上。",
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
