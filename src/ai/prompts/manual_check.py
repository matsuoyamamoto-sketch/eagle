"""マニュアルチェックリスト生成用プロンプト。

スタディ全体を 1 リクエストで処理し、AI には『記入漏れ』『整合性』のみ抽出させる。
『単位・桁数』『横断』は決定論的に生成側で展開する。
"""
from __future__ import annotations

import json

from ...parser.models import Sheet, Study


SYSTEM = """あなたは臨床試験の Data Manager (DM) です。SDV 前に DM が目視確認すべきチェックポイントを抽出してください。

# 出力カテゴリ (この 2 つのみ使用)
- 記入漏れ: EDC で必須化されていないが、条件付きで入力されているべき項目の入力漏れ確認
- 整合性: 単一フォーム内の項目間整合性 (日付ペアの前後・選択肢と他項目の連動・SAE 報告の整合 など)

# 重要なルール
- target_fields は **`ラベル(field名)` 形式** (例: `投与量(field3)`) で 1 件以上指定してください。
- 各チェックは **必ず sheet_name を 1 つ指定** し、その sheet_name に含まれる field のみを target_fields に使用してください (フォーム横断のチェックは別途決定論で生成するため不要です)。
- 上記 2 カテゴリ以外は出力しないでください。
- 自由記述 (text)・薬剤コーディング (drug)・MedDRA コーディング (meddra)・Note は対象外です。
- `has_default: true` の項目は記入漏れチェックの対象外です。
- 各シート 1〜3 件程度に絞り、汎用的すぎる注意喚起は避けてください。
- 1 つのチェックポイントには **target_fields を 1 件のみ** 指定してください (複数項目に関わる整合性は、主たる 1 項目を target に置き、文中で他項目を言及すること)。
- 出力は JSON スキーマ準拠で、**文法的に正しく、自然で正確な日本語**で記述してください。曖昧な機械翻訳調や文字化けした文字列は絶対に出力しないでください。
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


def build_study_prompt(study: Study, selected_sheet_names: list[str]) -> str:
    """スタディ全体を 1 リクエストにまとめるためのユーザープロンプト。"""
    sheets_payload = []
    selected = set(selected_sheet_names)
    for sheet in study.sheets:
        if sheet.name not in selected:
            continue
        items = _candidate_items(sheet)
        if not items:
            continue
        sheets_payload.append({"sheet_name": sheet.name, "items": items})
    payload = {"study_name": study.proper_name or study.name, "sheets": sheets_payload}
    return (
        "次の臨床試験のフォーム群について、DM が SDV 前に目視確認すべきチェックポイントを抽出してください。\n"
        "観点は『記入漏れ / 整合性』の 2 カテゴリのみです。各チェックは必ず sheet_name を指定し、\n"
        "その sheet_name に含まれる field のみを target_fields に使用してください。\n\n"
        f"```json\n{json.dumps(payload, ensure_ascii=False, indent=2)}\n```"
    )


# 後方互換 (テスト等で使われる場合)
def build_user_prompt(sheet: Sheet, study: Study | None = None) -> str:  # pragma: no cover
    items = _candidate_items(sheet)
    payload = {"sheet_name": sheet.name, "candidate_items": items}
    return f"```json\n{json.dumps(payload, ensure_ascii=False, indent=2)}\n```"


SCHEMA = {
    "type": "object",
    "properties": {
        "check_points": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "sheet_name": {"type": "string"},
                    "category": {"type": "string", "enum": ["記入漏れ", "整合性"]},
                    "target_fields": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "対象フィールドを `ラベル(field名)` 形式 (例: `投与量(field3)`) で 1 件以上。",
                    },
                    "check_point": {"type": "string"},
                    "rationale": {"type": "string"},
                    "severity": {"type": "string", "enum": ["high", "medium", "low"]},
                },
                "required": ["sheet_name", "category", "target_fields", "check_point", "severity"],
            },
        }
    },
    "required": ["check_points"],
}
