"""マニュアルチェックリスト生成用プロンプト (シート単位 AI 呼び出し)。"""
from __future__ import annotations

import json

from ...parser.models import Sheet


CATEGORIES = ["記入漏れ", "整合性"]


SYSTEM = """あなたは臨床試験の Data Manager (DM) です。SDV 前に DM が目視確認すべきチェックポイントを抽出してください。

# 出力カテゴリ (この 2 つのみ使用)
- 記入漏れ: EDC で必須化されていないが、条件付きで入力されているべき項目の入力漏れ確認
- 整合性: 単一フォーム内の項目間整合性 (日付ペアの前後・選択肢と他項目の連動・SAE 報告の整合 など)

# 重要なルール
- target_fields は **`ラベル(field名)` 形式** (例: `投与量(field3)`) で **1 件のみ** 指定してください。複数項目に関わる整合性は、主たる 1 項目を target に置き、文中で他項目を言及してください。
- target_fields に指定できるのは、candidate_items に含まれる field/label のみ。存在しないフィールドを推測・創作しないでください。
- フォーム横断のチェックは出力しないでください (このフォーム単独で完結するもののみ)。
- 自由記述 (text)・薬剤コーディング (drug)・MedDRA コーディング (meddra)・Note は対象外です (candidate_items に含まれません)。
- `has_default: true` の項目は記入漏れチェックの対象外です。
- 各シート 1〜5 件程度に絞り、汎用的すぎる注意喚起は避けてください。該当チェックがない場合は空配列を返してください。
- 出力は JSON スキーマ準拠で、**文法的に正しく、自然で正確な日本語**で記述してください。曖昧な機械翻訳調や文字化けした文字列は出力しないでください。
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
        "観点は『記入漏れ / 整合性』の 2 カテゴリのみ、フォーム単独で完結するチェックに限定します。\n\n"
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
                    "category": {"type": "string", "enum": CATEGORIES},
                    "target_fields": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "対象フィールドを `ラベル(field名)` 形式 (例: `投与量(field3)`) で 1 件指定。",
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
