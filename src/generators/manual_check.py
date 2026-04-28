"""マニュアルチェックリスト (Excel) 生成 — Cohere 利用。"""
# ... (インポート部分は変更なし) ...

def generate_check_points(
    study: Study,
    selected_sheet_names: list[str],
    client: CohereJSONClient | None = None,
    on_progress: Callable[[int, int, str], None] | None = None,
) -> list[dict]:
    target_sheets = [s for s in study.sheets if s.name in set(selected_sheet_names)]
    out: list[dict] = []
    total = len(target_sheets)
    client = client or CohereJSONClient()

    # 重複排除のためのセット
    seen_checkpoints = set()

    ai_skip = False
    consecutive_errors = 0
    for i, sheet in enumerate(target_sheets, start=1):
        if on_progress:
            on_progress(i, total, sheet.name)
        valid_fields = {fi.name for fi in sheet.field_items if fi.type != "FieldItem::Note"}

        if not ai_skip:
            try:
                data = client.chat_json(P.SYSTEM, P.build_user_prompt(sheet), P.SCHEMA)
                consecutive_errors = 0
                for cp in data.get("check_points", []):
                    category = cp.get("category", "")
                    if category not in P.CATEGORIES:
                        continue
                    
                    tgt_str = cp.get("target_field")
                    if not tgt_str:
                        continue
                    
                    check_text = cp.get("check_point", "")
                    
                    # 【追加】重複排除ロジック
                    # シート名、カテゴリ、ターゲット、チェック内容の組み合わせで一意性を確認
                    unique_key = (sheet.name, category, tgt_str, check_text)
                    if unique_key in seen_checkpoints:
                        continue
                    
                    referenced = re.findall(r"field\d+", tgt_str)
                    if not referenced or any(f not in valid_fields for f in referenced):
                        continue
                        
                    out.append(
                        {
                            "sheet": sheet.name,
                            "category": category,
                            "target_fields": [tgt_str],
                            "check_point": check_text,
                            "rationale": cp.get("rationale", ""),
                            "severity": cp.get("severity", "medium"),
                        }
                    )
                    seen_checkpoints.add(unique_key) # 登録済みとしてマーク
            except Exception as e:
                consecutive_errors += 1
                # ... (エラー処理は変更なし) ...
                if consecutive_errors >= 2:
                    ai_skip = True

        # 決定論的なチェックを追加
        for unit_check in _unit_digit_checks_for_sheet(sheet):
            # 決定論的なものも重複確認
            tgt_fields = unit_check["target_fields"][0]
            unique_key = (sheet.name, unit_check["category"], tgt_fields, unit_check["check_point"])
            if unique_key not in seen_checkpoints:
                out.append(unit_check)
                seen_checkpoints.add(unique_key)

    # ... (ソートと展開処理は変更なし) ...
    return _expand_one_target_per_row(out)

# ... (残りの関数は変更なし) ...