"""EAGLE — EDC Auto Generator & Logic Extractor (Streamlit UI)。"""
from __future__ import annotations

from datetime import datetime

import streamlit as st

from src.config import settings
from src.generators.edit_check import build_edit_check_workbook
from src.generators.manual_check import build_manual_check_workbook, generate_check_points
from src.generators.spec_excel import build_spec_workbook
from src.generators.test_scenario import build_test_scenario_workbook, generate_scenarios
from src.generators.validation_plan import build_validation_plan
from src.parser.edc_parser import load_study
from src.utils.zipper import docx_to_bytes, excel_to_bytes, files_to_zip_bytes

st.set_page_config(page_title="EAGLE", page_icon="📄", layout="centered")

# ---------------- ヘッダー ----------------
_EM = "color:#1F3864;font-weight:800;font-size:1.35em;"
st.markdown(
    f"""
<h3 style='margin-bottom:0.25rem;'>📄 EAGLE
<span style='font-size:0.75em;color:#475569;font-weight:500;margin-left:8px;'>
(<span style='{_EM}'>E</span>DC
<span style='{_EM}'>A</span>uto
<span style='{_EM}'>G</span>enerator &amp;
<span style='{_EM}'>L</span>ogic
<span style='{_EM}'>E</span>xtractor)
</span>
</h3>
""",
    unsafe_allow_html=True,
)
st.caption("EDC の設定ファイル (JSON) から、バリデーションプランやテストシナリオを自動生成します。")
st.divider()

# ---------------- サイドバー: AI設定 ----------------
with st.sidebar:
    st.markdown("### 🤖 AI 設定 (Cohere)")
    api_key_input = st.text_input(
        "API Key",
        value=settings.cohere_api_key,
        type="password",
        help="空欄の場合は .env の COHERE_API_KEY を使用",
    )
    model_input = st.text_input("Model", value=settings.cohere_model)
    rpm_input = st.number_input(
        "Requests / minute",
        min_value=1,
        max_value=120,
        value=settings.cohere_requests_per_minute,
        help="Trial キーは 20 以下を推奨",
    )

# ---------------- Step 1: JSON アップロード ----------------
st.markdown("##### ① 設定ファイル (JSON) のアップロード")
uploaded = st.file_uploader(
    "ここへファイルをドラッグ＆ドロップ または クリックして選択 (最大 200MB)",
    type=["json"],
    accept_multiple_files=False,
)

study = None
if uploaded is not None:
    try:
        study = load_study(uploaded.getvalue())
        st.success(
            f"✓ 読み込み成功: **{study.name}**  "
            f"(フォーム {len(study.sheets)} / 項目 {study.total_field_items():,})"
        )
    except Exception as e:
        st.error(f"JSON の読込に失敗しました: {e}")

# ---------------- Step 2: ドキュメント選択 ----------------
st.markdown("##### ② 生成するドキュメントの選択")

DOCS = [
    ("spec",     "EDC仕様書 (Excel)",            False, False),
    ("vplan",    "バリデーションプラン (Word)",   False, True),
    ("echeck",   "エディットチェック確認書 (Excel)", False, False),
    ("scenario", "テストシナリオ (Excel)",        True,  True),
    ("manual",   "マニュアルチェックリスト (Excel)", True, True),
]

cols = st.columns(2)
selections: dict[str, bool] = {}
for i, (key, label, is_ai, default) in enumerate(DOCS):
    with cols[i % 2]:
        suffix = "  🤖 AI生成" if is_ai else ""
        selections[key] = st.checkbox(label + suffix, value=default, key=f"chk_{key}")

# ---------------- Step 2.5: AI 対象フォーム選択 ----------------
ai_targets: list[str] = []
ai_selected = selections.get("scenario") or selections.get("manual")
if ai_selected and study is not None:
    st.markdown("##### ②' AI 生成対象のフォーム選択")
    st.caption(
        f"⚠️ Trial キーはレート制限が厳しいため、最初は 3〜5 件で動作確認することを推奨します "
        f"(全{len(study.sheets)}フォーム)。"
    )
    sheet_names = [s.name for s in study.sheets]
    default_targets = sheet_names[: min(3, len(sheet_names))]
    ai_targets = st.multiselect(
        "対象フォーム",
        options=sheet_names,
        default=default_targets,
        key="ai_targets",
    )

# ---------------- Step 3: 実行 ----------------
st.markdown("##### ③ 実行")
run = st.button("⚙️ ドキュメントを生成する", type="primary", use_container_width=True)


def _make_client():
    from src.ai.cohere_client import CohereJSONClient

    return CohereJSONClient(api_key=api_key_input or None, model=model_input, rpm=int(rpm_input))


if run:
    if study is None:
        st.warning("先に JSON ファイルをアップロードしてください。")
        st.stop()
    if not any(selections.values()):
        st.warning("生成するドキュメントを 1 つ以上選択してください。")
        st.stop()
    if ai_selected and not ai_targets:
        st.warning("AI 生成対象のフォームを 1 つ以上選択してください。")
        st.stop()

    files: dict[str, bytes] = {}
    selected_keys = [k for k, v in selections.items() if v]
    progress = st.progress(0.0, text="準備中…")
    detail = st.empty()
    step_total = len(selected_keys)
    step = {"n": 0}

    def tick(label: str):
        step["n"] += 1
        progress.progress(step["n"] / step_total, text=label)
        detail.empty()

    try:
        if selections.get("spec"):
            tick("EDC仕様書 (Excel) を生成中…")
            wb = build_spec_workbook(study)
            files[f"{study.name}_EDC仕様書.xlsx"] = excel_to_bytes(wb)

        if selections.get("vplan"):
            tick("バリデーションプラン (Word) を生成中…")
            doc = build_validation_plan(study)
            files[f"{study.name}_バリデーションプラン.docx"] = docx_to_bytes(doc)

        if selections.get("echeck"):
            tick("エディットチェック確認書 (Excel) を生成中…")
            wb, _ = build_edit_check_workbook(study)
            files[f"{study.name}_エディットチェック確認書.xlsx"] = excel_to_bytes(wb)

        # ----- AI: テストシナリオ -----
        if selections.get("scenario"):
            tick(f"テストシナリオ (AI, {len(ai_targets)}フォーム) を生成中…")
            try:
                client = _make_client()
                sub_bar = st.progress(0.0, text="AI 呼び出し中…")

                def _on_p(i: int, total: int, name: str):
                    sub_bar.progress(i / total, text=f"テストシナリオ {i}/{total}: {name}")

                scenarios = generate_scenarios(study, ai_targets, client, _on_p)
                wb = build_test_scenario_workbook(study, scenarios)
                files[f"{study.name}_テストシナリオ.xlsx"] = excel_to_bytes(wb)
                sub_bar.empty()
            except Exception as e:
                st.error(f"テストシナリオ生成エラー: {e}")

        # ----- AI: マニュアルチェック -----
        if selections.get("manual"):
            tick(f"マニュアルチェックリスト (AI, {len(ai_targets)}フォーム) を生成中…")
            try:
                client = _make_client()
                sub_bar = st.progress(0.0, text="AI 呼び出し中…")

                def _on_p(i: int, total: int, name: str):
                    sub_bar.progress(i / total, text=f"マニュアルチェック {i}/{total}: {name}")

                points = generate_check_points(study, ai_targets, client, _on_p)
                wb = build_manual_check_workbook(study, points)
                files[f"{study.name}_マニュアルチェックリスト.xlsx"] = excel_to_bytes(wb)
                sub_bar.empty()
            except Exception as e:
                st.error(f"マニュアルチェックリスト生成エラー: {e}")

        progress.progress(1.0, text="完了")

        if not files:
            st.warning("生成可能なドキュメントがありませんでした。")
            st.stop()

        zip_bytes = files_to_zip_bytes(files)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        st.success(f"✓ {len(files)} 件のドキュメントを生成しました。")
        st.download_button(
            label="📦 ZIP でダウンロード",
            data=zip_bytes,
            file_name=f"{study.name}_EDC_docs_{ts}.zip",
            mime="application/zip",
            use_container_width=True,
        )
    except Exception as e:
        st.error(f"生成中にエラーが発生しました: {e}")
        st.exception(e)
