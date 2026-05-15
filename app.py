"""EAGLE — EDC Auto Generator & Logic Extractor (Streamlit UI)。"""
from __future__ import annotations

from datetime import datetime

import streamlit as st
from streamlit_local_storage import LocalStorage

from src.config import settings
from src.generators.edit_check import build_edit_check_workbook
from src.generators.manual_check import build_manual_check_workbook, generate_check_points
from src.generators.spec_excel import build_spec_workbook
from src.generators.test_scenario import build_test_scenario_workbook, generate_scenarios
from src.generators.validation_plan import build_validation_plan
from src.parser.edc_parser import load_study
from src.utils.zipper import docx_to_bytes, excel_to_bytes, files_to_zip_bytes

st.set_page_config(page_title="EAGLE", page_icon="🦅", layout="wide")

# ---------------- カラーテーマ ----------------
NAVY = "#0F1F3D"
NAVY_2 = "#1F3864"
GOLD = "#C9A24C"
GOLD_LIGHT = "#F5E9CC"
SLATE_50 = "#F8FAFC"
SLATE_200 = "#E2E8F0"
SLATE_500 = "#64748B"
SLATE_700 = "#334155"

# ---------------- グローバルスタイル ----------------
st.markdown(
    f"""
<style>
  .block-container {{ padding-top: 1.2rem; max-width: 1280px; }}
  /* ヒーローバー */
  .eagle-hero {{
    display:flex; align-items:center; gap:18px;
    background: linear-gradient(135deg, {NAVY} 0%, {NAVY_2} 100%);
    color:#fff; padding:18px 24px; border-radius:14px;
    box-shadow: 0 4px 14px rgba(15,31,61,0.18);
    border: 1px solid rgba(201,162,76,0.25);
    margin-bottom: 18px;
  }}
  .eagle-hero .logo-badge {{
    width:56px; height:56px; border-radius:14px;
    background: rgba(255,255,255,0.06);
    border: 1px solid rgba(201,162,76,0.4);
    display:flex; align-items:center; justify-content:center;
    color: {GOLD};
    flex-shrink:0;
  }}
  .eagle-hero h1 {{
    margin:0; font-size: 1.7rem; font-weight:900;
    letter-spacing: 0.18em; color:#fff;
  }}
  .eagle-hero .subtitle {{
    font-size: 0.78rem; color: {GOLD_LIGHT};
    letter-spacing: 0.08em; font-weight:600; text-transform:uppercase;
    margin-top: 2px;
  }}
  .eagle-hero .desc {{
    font-size: 0.78rem; color: rgba(255,255,255,0.72);
    margin-top: 4px;
  }}

  /* セクション見出し */
  .step-head {{
    display:flex; align-items:center; gap:10px;
    margin: 6px 0 10px;
  }}
  .step-num {{
    width:28px; height:28px; border-radius:50%;
    background:{NAVY}; color:{GOLD};
    display:flex; align-items:center; justify-content:center;
    font-weight:800; font-size:0.85rem;
    border: 2px solid {GOLD};
  }}
  .step-title {{
    font-weight:700; color:{NAVY}; font-size:1.0rem;
  }}

  /* サマリーカード */
  .summary-card {{
    background:#fff; border:1px solid {SLATE_200}; border-radius:12px;
    padding:18px; box-shadow: 0 1px 2px rgba(0,0,0,0.03);
  }}
  .summary-card .study-name {{
    font-weight:800; color:{NAVY}; font-size:1.05rem;
    border-left: 4px solid {GOLD}; padding-left:10px; margin-bottom:14px;
    line-height:1.4;
  }}
  .kpi-grid {{ display:grid; grid-template-columns: repeat(2, 1fr); gap:10px; }}
  .kpi {{
    background:{SLATE_50}; border:1px solid {SLATE_200}; border-radius:10px;
    padding:10px 12px; text-align:center;
  }}
  .kpi .label {{ font-size:0.7rem; color:{SLATE_500}; margin-bottom:2px; }}
  .kpi .value {{ font-size:1.4rem; font-weight:800; color:{NAVY}; }}
  .kpi .value.gold {{ color:{GOLD}; }}
  .kpi .sub {{ font-size:0.7rem; color:{SLATE_500}; }}

  .placeholder-card {{
    background:{SLATE_50}; border:2px dashed {SLATE_200}; border-radius:12px;
    padding:36px 18px; text-align:center; color:{SLATE_500};
  }}
  .placeholder-card .ico {{ color:{NAVY_2}; opacity:0.35; margin-bottom:8px; }}

  /* ドキュメントカード — st.container(border) を装飾 */
  div[data-testid="stVerticalBlockBorderWrapper"]:has(.doc-card-marker) {{
    border-radius: 12px !important;
    border: 1px solid {SLATE_200} !important;
    background: #fff;
    transition: border-color 0.15s, box-shadow 0.15s, background 0.15s;
  }}
  div[data-testid="stVerticalBlockBorderWrapper"]:has(.doc-card-marker:hover) {{
    border-color: {GOLD} !important;
    box-shadow: 0 4px 12px rgba(201,162,76,0.15);
  }}
  /* 選択中カード — ゴールド枠太め + 薄いゴールド地 */
  div[data-testid="stVerticalBlockBorderWrapper"]:has(.doc-card-marker.selected) {{
    border: 2px solid {GOLD} !important;
    background: linear-gradient(180deg, #FFFCF5 0%, #FFFFFF 80%) !important;
    box-shadow: 0 4px 14px rgba(201,162,76,0.22) !important;
  }}
  /* カード内のトグル右隣に置く自前ラベル (トグルのスイッチと縦中央揃え) */
  .card-toggle-label {{
    font-size: 0.78rem;
    color: {SLATE_700};
    font-weight: 500;
    line-height: 1.4;
    user-select: none;
    display: flex;
    align-items: center;
    height: 100%;
    min-height: 28px;
    padding-top: 12px;
  }}
  .card-toggle-label.selected {{
    color: {NAVY};
    font-weight: 700;
  }}
  .doc-card-head {{ display:flex; align-items:center; gap:10px; margin-bottom:6px; }}
  .doc-card-head .ico {{
    width:34px; height:34px; border-radius:9px;
    background:{SLATE_50}; color:{NAVY};
    display:flex; align-items:center; justify-content:center; flex-shrink:0;
  }}
  .doc-card-head .title {{ font-weight:700; color:{NAVY}; font-size:0.93rem; }}
  .doc-card-head .ai-badge {{
    margin-left:auto;
    background:{GOLD_LIGHT}; color:#7a5c1f;
    border:1px solid {GOLD}; border-radius:999px;
    font-size:0.65rem; font-weight:700; padding:2px 8px;
    letter-spacing:0.05em;
  }}
  .doc-card-desc {{ font-size:0.75rem; color:{SLATE_500}; line-height:1.5; margin-bottom:6px; }}
  .doc-card-meta {{
    font-size:0.68rem; color:{SLATE_500};
    background:{SLATE_50}; border-radius:6px; padding:3px 8px;
    display:inline-block;
  }}

  /* 結果ファイルカード */
  .result-file {{
    display:flex; align-items:center; gap:12px;
    background:#fff; border:1px solid {SLATE_200}; border-radius:10px;
    padding:10px 14px;
  }}
  .result-file .ico {{
    width:32px; height:32px; border-radius:8px;
    background:{GOLD_LIGHT}; color:{NAVY};
    display:flex; align-items:center; justify-content:center; flex-shrink:0;
  }}
  .result-file .name {{ font-weight:600; color:{NAVY}; font-size:0.85rem; }}
  .result-file .size {{ font-size:0.7rem; color:{SLATE_500}; }}

  /* プライマリボタンを Gold に */
  .stButton > button[kind="primary"] {{
    background: linear-gradient(135deg, {NAVY} 0%, {NAVY_2} 100%) !important;
    border: 1px solid {GOLD} !important;
    color: #fff !important;
    font-weight:700; letter-spacing:0.05em;
  }}
  .stButton > button[kind="primary"]:hover {{
    box-shadow: 0 4px 14px rgba(201,162,76,0.35) !important;
  }}

  /* セクション仕切り */
  hr {{ margin: 0.8rem 0 !important; }}
</style>
""",
    unsafe_allow_html=True,
)

# ---------------- SVG アイコン ----------------
EAGLE_SVG = """
<svg viewBox="0 0 64 64" fill="none" xmlns="http://www.w3.org/2000/svg" width="36" height="36">
  <!-- 翼を広げた鷲のシルエット -->
  <path d="M32 14 L29 22 C24 20 18 19 10 22 C16 24 20 27 24 30 C18 31 13 33 8 37 C16 36 22 37 26 38 C22 41 19 44 16 49 C22 46 27 44 30 43 L31 51 L33 51 L34 43 C37 44 42 46 48 49 C45 44 42 41 38 38 C42 37 48 36 56 37 C51 33 46 31 40 30 C44 27 48 24 54 22 C46 19 40 20 35 22 L32 14 Z"
        fill="currentColor"/>
  <circle cx="32" cy="18" r="2.2" fill="#0F1F3D"/>
</svg>
"""

ICON_UPLOAD = """<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" width="18" height="18"><path stroke-linecap="round" stroke-linejoin="round" d="M3 16.5v2.25A2.25 2.25 0 0 0 5.25 21h13.5A2.25 2.25 0 0 0 21 18.75V16.5m-13.5-9L12 3m0 0 4.5 4.5M12 3v13.5"/></svg>"""
ICON_DOCS = """<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" width="18" height="18"><path stroke-linecap="round" stroke-linejoin="round" d="M9 12h6m-6 4h6m2 5H7a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5.586a1 1 0 0 1 .707.293l5.414 5.414a1 1 0 0 1 .293.707V19a2 2 0 0 1-2 2Z"/></svg>"""
ICON_RUN = """<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" width="18" height="18"><path stroke-linecap="round" stroke-linejoin="round" d="M9.594 3.94c.09-.542.56-.94 1.11-.94h2.593c.55 0 1.02.398 1.11.94l.213 1.281c.063.374.313.686.645.87.074.04.147.083.22.127.325.196.72.257 1.075.124l1.217-.456a1.125 1.125 0 0 1 1.37.49l1.296 2.247a1.125 1.125 0 0 1-.26 1.431l-1.003.827c-.293.241-.438.613-.43.992a7.723 7.723 0 0 1 0 .255c-.008.378.137.75.43.991l1.004.827c.424.35.534.955.26 1.43l-1.298 2.247a1.125 1.125 0 0 1-1.369.491l-1.217-.456c-.355-.133-.75-.072-1.076.124a6.47 6.47 0 0 1-.22.128c-.331.183-.581.495-.644.869l-.213 1.281c-.09.543-.56.94-1.11.94h-2.594c-.55 0-1.019-.398-1.11-.94l-.213-1.281c-.062-.374-.312-.686-.644-.87a6.52 6.52 0 0 1-.22-.127c-.325-.196-.72-.257-1.076-.124l-1.217.456a1.125 1.125 0 0 1-1.369-.49l-1.297-2.247a1.125 1.125 0 0 1 .26-1.431l1.004-.827c.292-.24.437-.613.43-.991a6.932 6.932 0 0 1 0-.255c.007-.38-.138-.751-.43-.992l-1.004-.827a1.125 1.125 0 0 1-.26-1.43l1.297-2.247a1.125 1.125 0 0 1 1.37-.491l1.216.456c.356.133.751.072 1.076-.124.072-.044.146-.087.22-.128.332-.183.582-.495.644-.869l.214-1.28Z"/><path stroke-linecap="round" stroke-linejoin="round" d="M15 12a3 3 0 1 1-6 0 3 3 0 0 1 6 0Z"/></svg>"""
ICON_EXCEL = """<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" width="18" height="18"><rect x="3.25" y="4.25" width="17.5" height="15.5" rx="1.5" stroke-linejoin="round"/><path stroke-linecap="round" stroke-linejoin="round" d="M3.25 9.5h17.5M3.25 14.5h17.5M9 4.25v15.5M15 4.25v15.5"/></svg>"""
ICON_WORD = """<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" width="18" height="18"><path stroke-linecap="round" stroke-linejoin="round" d="M19.5 14.25v-2.625a3.375 3.375 0 0 0-3.375-3.375h-1.5A1.125 1.125 0 0 1 13.5 7.125v-1.5A3.375 3.375 0 0 0 10.125 2.25H8.25M10.5 2.25H5.625c-.621 0-1.125.504-1.125 1.125v17.25c0 .621.504 1.125 1.125 1.125h12.75c.621 0 1.125-.504 1.125-1.125V11.25a9 9 0 0 0-9-9Z"/></svg>"""
ICON_AI = """<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" width="18" height="18"><path stroke-linecap="round" stroke-linejoin="round" d="M9.813 15.904 9 18.75l-.813-2.846a4.5 4.5 0 0 0-3.09-3.09L2.25 12l2.846-.813a4.5 4.5 0 0 0 3.09-3.09L9 5.25l.813 2.846a4.5 4.5 0 0 0 3.09 3.09L15.75 12l-2.846.813a4.5 4.5 0 0 0-3.09 3.09ZM18.259 8.715 18 9.75l-.259-1.035a3.375 3.375 0 0 0-2.455-2.456L14.25 6l1.036-.259a3.375 3.375 0 0 0 2.455-2.456L18 2.25l.259 1.035a3.375 3.375 0 0 0 2.456 2.456L21.75 6l-1.035.259a3.375 3.375 0 0 0-2.456 2.456Z"/></svg>"""
ICON_DOWNLOAD = """<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" width="18" height="18"><path stroke-linecap="round" stroke-linejoin="round" d="M3 16.5v2.25A2.25 2.25 0 0 0 5.25 21h13.5A2.25 2.25 0 0 0 21 18.75V16.5M16.5 12 12 16.5m0 0L7.5 12m4.5 4.5V3"/></svg>"""

# ---------------- ヒーローバー ----------------
st.markdown(
    f"""
<div class="eagle-hero">
  <div class="logo-badge">{EAGLE_SVG}</div>
  <div style="flex:1;">
    <h1>EAGLE</h1>
    <div class="subtitle">EDC Auto Generator &amp; Logic Extractor</div>
    <div class="desc">EDC の設定ファイル (JSON) から、バリデーションプランやテストシナリオを自動生成します。</div>
  </div>
</div>
""",
    unsafe_allow_html=True,
)

# ---------------- サイドバー: AI設定 ----------------
_LS_KEY = "eagle_cohere_api_key"
_local_storage = LocalStorage()

with st.sidebar:
    st.markdown(
        f"<div style='display:flex;align-items:center;gap:8px;color:{NAVY};font-weight:700;font-size:1rem;margin-bottom:10px;'>"
        f"<span style='color:{GOLD};'>{ICON_AI}</span> AI 設定 (Cohere)</div>",
        unsafe_allow_html=True,
    )

    saved_key = _local_storage.getItem(_LS_KEY) or ""
    effective_key = saved_key or settings.cohere_api_key or ""

    # ステータスバッジ
    if effective_key:
        masked = effective_key[:4] + "•" * 8 + effective_key[-2:] if len(effective_key) >= 6 else "••••"
        src = "ブラウザ保存" if saved_key else "環境変数"
        st.markdown(
            f"""
<div style="background:#ECFDF5;border:1px solid #86EFAC;border-radius:10px;
            padding:10px 12px;margin-bottom:12px;">
  <div style="font-size:0.8rem;font-weight:700;color:#15803D;display:flex;align-items:center;gap:6px;">
    ✓ 接続準備 OK
    <span style="margin-left:auto;font-size:0.65rem;background:#fff;color:#166534;
                 border:1px solid #86EFAC;border-radius:999px;padding:1px 8px;font-weight:600;">{src}</span>
  </div>
  <div style="font-family:monospace;font-size:0.78rem;color:{SLATE_700};margin-top:4px;">{masked}</div>
</div>
""",
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            f"""
<div style="background:#FFFBEB;border:1px solid #FCD34D;border-radius:10px;
            padding:10px 12px;margin-bottom:12px;">
  <div style="font-size:0.8rem;font-weight:700;color:#92400E;">⚠ API キー未設定</div>
  <div style="font-size:0.72rem;color:#78350F;margin-top:3px;line-height:1.4;">
    AI 生成 (テストシナリオ / マニュアルチェック) を使用するには、下のフォームにキーを入力してください。
  </div>
</div>
""",
            unsafe_allow_html=True,
        )

    with st.expander("キーを編集 / 保存", expanded=not effective_key):
        api_key_input = st.text_input(
            "API Key",
            value=saved_key or settings.cohere_api_key,
            type="password",
            help="「ブラウザに保存」を押すとこの端末のブラウザに記憶されます (localStorage)",
            label_visibility="collapsed",
            placeholder="Cohere API Key を貼り付け…",
        )
        c1, c2 = st.columns([2, 1])
        with c1:
            if st.button(
                "💾 ブラウザに保存",
                use_container_width=True,
                disabled=not api_key_input or api_key_input == saved_key,
            ):
                _local_storage.setItem(_LS_KEY, api_key_input)
                st.success("保存しました", icon="✅")
                st.rerun()
        with c2:
            if st.button("🗑️ 削除", use_container_width=True, disabled=not saved_key):
                _local_storage.deleteItem(_LS_KEY)
                st.info("削除しました", icon="ℹ️")
                st.rerun()

    if not 'api_key_input' in dir():
        api_key_input = effective_key

    st.markdown("<div style='margin-top:8px;'></div>", unsafe_allow_html=True)

    # ----- モデル選択 -----
    # (id, 表示名, 速度1-5, 精度1-5, コスト相対, 説明)
    MODEL_CATALOG: list[tuple[str, str, int, int, str, str]] = [
        ("command-a-03-2025",      "Command A (Flagship)", 2, 5, "高",   "最高精度・複雑な推論。Trial だとレート制限が厳しめ。"),
        ("command-r-plus-08-2024", "Command R+",           3, 4, "中高", "高品質と速度のバランス型。EAGLE のデフォルト。"),
        ("command-r-08-2024",      "Command R",            4, 3, "中",   "標準。実用的な速度と精度。"),
        ("command-r7b-12-2024",    "Command R7B",          5, 2, "低",   "最速・最安。短い JSON 抽出向け。動作確認に最適。"),
    ]
    model_ids = [m[0] for m in MODEL_CATALOG]
    current_model = settings.cohere_model
    if current_model not in model_ids:
        MODEL_CATALOG.append((current_model, current_model, 0, 0, "?", "(カスタムモデル)"))
        model_ids.append(current_model)

    def _bars(n: int, color: str) -> str:
        full = "●" * n
        empty = "○" * (5 - n)
        return f"<span style='color:{color};letter-spacing:1px;'>{full}</span><span style='color:#cbd5e1;letter-spacing:1px;'>{empty}</span>"

    st.markdown(
        f"<div style='font-size:0.78rem;color:{NAVY};font-weight:700;margin-bottom:4px;'>Model</div>",
        unsafe_allow_html=True,
    )
    selected_idx = model_ids.index(current_model) if current_model in model_ids else 0
    model_input = st.selectbox(
        "Model",
        options=model_ids,
        index=selected_idx,
        format_func=lambda mid: next((m[1] for m in MODEL_CATALOG if m[0] == mid), mid),
        label_visibility="collapsed",
    )

    # 選択中モデルのスペック表示
    spec = next((m for m in MODEL_CATALOG if m[0] == model_input), None)
    if spec:
        _, label, speed, quality, cost, desc = spec
        st.markdown(
            f"""
<div style="background:{SLATE_50};border:1px solid {SLATE_200};border-radius:10px;
            padding:10px 12px;margin-top:6px;font-size:0.75rem;">
  <div style="font-family:monospace;color:{SLATE_500};font-size:0.7rem;margin-bottom:6px;">{model_input}</div>
  <div style="display:grid;grid-template-columns:auto 1fr;gap:4px 10px;align-items:center;">
    <span style="color:{SLATE_700};">速度</span><span>{_bars(speed, '#16a34a') if speed else '-'}</span>
    <span style="color:{SLATE_700};">精度</span><span>{_bars(quality, '#2563eb') if quality else '-'}</span>
    <span style="color:{SLATE_700};">コスト</span><span style="color:{NAVY};font-weight:600;">{cost}</span>
  </div>
  <div style="color:{SLATE_500};margin-top:8px;line-height:1.45;">{desc}</div>
</div>
""",
            unsafe_allow_html=True,
        )

    # 全モデル比較 (折りたたみ)
    with st.expander("モデル比較表を見る"):
        rows_html = ""
        for mid, label, speed, quality, cost, _desc in MODEL_CATALOG:
            highlight = f"background:{GOLD_LIGHT};" if mid == model_input else ""
            rows_html += (
                f"<tr style='{highlight}'>"
                f"<td style='padding:4px 6px;font-weight:600;color:{NAVY};'>{label}</td>"
                f"<td style='padding:4px 6px;'>{_bars(speed, '#16a34a') if speed else '-'}</td>"
                f"<td style='padding:4px 6px;'>{_bars(quality, '#2563eb') if quality else '-'}</td>"
                f"<td style='padding:4px 6px;color:{SLATE_700};'>{cost}</td>"
                f"</tr>"
            )
        st.markdown(
            f"""
<table style="width:100%;font-size:0.7rem;border-collapse:collapse;">
  <thead>
    <tr style="border-bottom:1px solid {SLATE_200};color:{SLATE_500};text-align:left;">
      <th style="padding:4px 6px;">モデル</th><th style="padding:4px 6px;">速度</th>
      <th style="padding:4px 6px;">精度</th><th style="padding:4px 6px;">コスト</th>
    </tr>
  </thead>
  <tbody>{rows_html}</tbody>
</table>
<div style="font-size:0.68rem;color:{SLATE_500};margin-top:6px;line-height:1.5;">
  💡 <b>動作確認は R7B</b>、<b>本番は R+ または A</b> がおすすめ。<br>
  Trial キーは月の使用量とレート制限に注意してください。
</div>
""",
            unsafe_allow_html=True,
        )

    rpm_input = st.number_input(
        "Requests / minute",
        min_value=1,
        max_value=120,
        value=settings.cohere_requests_per_minute,
        help="Trial キーは 20 以下を推奨",
    )

# ---------------- メイン: 2 カラム ----------------
left, right = st.columns([1.4, 1.0], gap="large")

# === LEFT: ステップフォーム ===
with left:
    # Step ①
    st.markdown(
        f'<div class="step-head"><div class="step-num">1</div>'
        f'<div class="step-title">設定ファイル (JSON) のアップロード</div></div>',
        unsafe_allow_html=True,
    )
    uploaded = st.file_uploader(
        "ここへファイルをドラッグ＆ドロップ または クリックして選択 (最大 200MB)",
        type=["json"],
        accept_multiple_files=False,
        label_visibility="collapsed",
    )

    study = None
    if uploaded is not None:
        try:
            study = load_study(uploaded.getvalue())
        except Exception as e:
            st.error(f"JSON の読込に失敗しました: {e}")

    # Step ②
    st.markdown(
        f'<div class="step-head" style="margin-top:18px;"><div class="step-num">2</div>'
        f'<div class="step-title">生成するドキュメントの選択</div></div>',
        unsafe_allow_html=True,
    )

    DOCS = [
        ("spec",     "EDC仕様書",                 "Excel", ICON_EXCEL, False, False,
         "全フォーム・全項目の定義をシート別に書き出した網羅仕様書。"),
        ("vplan",    "バリデーションプラン",       "Word",  ICON_WORD,  False, False,
         "EDC バリデーション戦略のドラフト文書。試験情報を自動充填。"),
        ("echeck",   "エディットチェック確認書",   "Excel", ICON_EXCEL, False, False,
         "JSON 内の validators を抽出し、確認用一覧として出力。"),
        ("scenario", "テストシナリオ",             "Excel", ICON_EXCEL, True,  False,
         "AI が各フォームの入力テストケースを生成。Trial キーは件数を絞ってください。"),
        ("manual",   "マニュアルチェックリスト",   "Excel", ICON_EXCEL, True,  False,
         "AI が項目間の論理矛盾チェック観点を抽出。"),
    ]

    selections: dict[str, bool] = {}
    grid_cols = st.columns(2)
    for i, (key, label, fmt, icon, is_ai, default, desc) in enumerate(DOCS):
        with grid_cols[i % 2]:
            sel_key = f"chk_{key}"
            if sel_key not in st.session_state:
                st.session_state[sel_key] = default
            is_selected = bool(st.session_state[sel_key])
            with st.container(border=True):
                marker_cls = "doc-card-marker selected" if is_selected else "doc-card-marker"
                st.markdown(f'<span class="{marker_cls}"></span>', unsafe_allow_html=True)
                ai_badge = '<span class="ai-badge">AI</span>' if is_ai else ''
                st.markdown(
                    f'<div class="doc-card-head">'
                    f'<div class="ico">{icon}</div>'
                    f'<div class="title">{label}</div>'
                    f'{ai_badge}</div>'
                    f'<div class="doc-card-desc">{desc}</div>'
                    f'<div class="doc-card-meta">{fmt}</div>',
                    unsafe_allow_html=True,
                )
                tcol1, tcol2 = st.columns([1, 5])
                with tcol1:
                    selections[key] = st.toggle(
                        "選択",
                        key=sel_key,
                        label_visibility="collapsed",
                    )
                with tcol2:
                    label_cls = "card-toggle-label selected" if is_selected else "card-toggle-label"
                    st.markdown(
                        f'<div class="{label_cls}">このドキュメントを生成する</div>',
                        unsafe_allow_html=True,
                    )

    # Step ②'
    ai_targets: list[str] = []
    ai_selected = selections.get("scenario") or selections.get("manual")
    if ai_selected and study is not None:
        st.markdown(
            f'<div class="step-head" style="margin-top:18px;">'
            f'<div class="step-num" style="background:{GOLD};color:{NAVY};border-color:{NAVY};">★</div>'
            f'<div class="step-title">AI 生成対象のフォーム選択</div></div>',
            unsafe_allow_html=True,
        )
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
            label_visibility="collapsed",
        )

    # Step ③
    st.markdown(
        f'<div class="step-head" style="margin-top:18px;"><div class="step-num">3</div>'
        f'<div class="step-title">実行</div></div>',
        unsafe_allow_html=True,
    )
    run = st.button("⚙️  ドキュメントを生成する", type="primary", use_container_width=True)

# === RIGHT: Study サマリー ===
with right:
    st.markdown(
        f'<div class="step-head"><div class="step-num" style="background:{GOLD};color:{NAVY};border-color:{NAVY};">i</div>'
        f'<div class="step-title">試験情報プレビュー</div></div>',
        unsafe_allow_html=True,
    )

    if study is not None:
        arms = getattr(study, "sheet_groups", None) or []
        arms_count = len(arms) if arms else 0
        try:
            total_items = study.total_field_items()
        except Exception:
            total_items = 0
        st.markdown(
            f"""
<div class="summary-card">
  <div class="study-name">{study.name}</div>
  <div class="kpi-grid">
    <div class="kpi"><div class="label">総フォーム数</div><div class="value gold">{len(study.sheets)}</div></div>
    <div class="kpi"><div class="label">総項目数</div><div class="value gold">{total_items:,}</div></div>
    <div class="kpi"><div class="label">割付グループ</div><div class="value">{arms_count}</div><div class="sub">Arms</div></div>
    <div class="kpi"><div class="label">読込状態</div><div class="value" style="font-size:1rem;color:#15803d;">✓ OK</div></div>
  </div>
</div>
""",
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            f"""
<div class="placeholder-card">
  <div class="ico">{EAGLE_SVG}</div>
  <div style="font-weight:600;color:{SLATE_700};margin-bottom:4px;">JSON 未読み込み</div>
  <div style="font-size:0.8rem;">左の ① からファイルをアップロードすると、<br>試験概要をここに表示します。</div>
</div>
""",
            unsafe_allow_html=True,
        )


def _make_client():
    from src.ai.cohere_client import CohereJSONClient

    return CohereJSONClient(api_key=api_key_input or None, model=model_input, rpm=int(rpm_input))


# ---------------- 実行処理 ----------------
def _file_icon(name: str) -> str:
    return ICON_WORD if name.lower().endswith(".docx") else ICON_EXCEL


def _mime_for(name: str) -> str:
    n = name.lower()
    if n.endswith(".docx"):
        return "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    if n.endswith(".xlsm"):
        return "application/vnd.ms-excel.sheet.macroEnabled.12"
    return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def _fmt_size(n: int) -> str:
    if n < 1024:
        return f"{n} B"
    if n < 1024 * 1024:
        return f"{n / 1024:.1f} KB"
    return f"{n / (1024 * 1024):.2f} MB"


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

    # 前回結果をクリア
    st.session_state.pop("generated_files", None)
    st.session_state.pop("generated_study_name", None)
    st.session_state.pop("generated_ts", None)

    files: dict[str, bytes] = {}
    errors: list[tuple[str, str]] = []
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

        def _make_ai_runner(label: str):
            """st.status パネル + イベントハンドラを生成。"""
            import time as _time
            status = st.status(f"🤖 {label} 準備中…", expanded=True)
            sub_bar = status.progress(0.0, text="待機中…")
            state = {"start": _time.monotonic(), "elapsed_per_sheet": [], "current": ""}

            def _ts() -> str:
                return datetime.now().strftime("%H:%M:%S")

            def _on_progress(i: int, total: int, name: str):
                state["current"] = name
                avg = (sum(state["elapsed_per_sheet"]) / len(state["elapsed_per_sheet"])
                       if state["elapsed_per_sheet"] else 0)
                eta = avg * (total - i + 1)
                eta_str = f" / 残り目安 {eta:.0f}秒" if avg > 0 else ""
                sub_bar.progress((i - 1) / total,
                                 text=f"{i}/{total}: {name}{eta_str}")
                status.update(label=f"🤖 {label} ({i}/{total}) — {name}")

            def _on_event(ev: dict):
                ph = ev.get("phase")
                if ph == "sheet_start":
                    skip_note = " (AI スキップ中)" if ev.get("skipped") else ""
                    status.write(f"`[{_ts()}]` ▶ **{ev['name']}** 開始 ({ev['i']}/{ev['total']}){skip_note}")
                elif ph == "rate_wait":
                    status.write(f"`[{_ts()}]` ⏸ レート制限待機 {ev['wait']:.1f}秒…")
                elif ph == "request_start":
                    status.write(f"`[{_ts()}]` 　→ Cohere に送信中…")
                elif ph == "heartbeat":
                    status.write(f"`[{_ts()}]` 　… 応答待ち {ev['elapsed']:.0f}秒経過{ev.get('note', '')}")
                elif ph == "request_end":
                    status.write(f"`[{_ts()}]` 　← 応答受信 ({ev['elapsed']:.1f}秒, {ev['chars']:,}文字)")
                elif ph == "request_error":
                    status.write(f"`[{_ts()}]` ⚠ 通信エラー ({ev['elapsed']:.1f}秒): `{ev['error'][:120]}`")
                elif ph == "retry":
                    status.write(f"`[{_ts()}]` 🔁 リトライ {ev['attempt']}/{ev['max']} — {ev['wait']:.0f}秒待機")
                elif ph == "sheet_end":
                    state["elapsed_per_sheet"].append(ev["elapsed"])
                    status.write(f"`[{_ts()}]` ✓ {ev['name']} 完了 ({ev['elapsed']:.1f}秒, {ev.get('items', 0)}件抽出)")
                elif ph == "sheet_error":
                    state["elapsed_per_sheet"].append(ev["elapsed"])
                    note = " — 以降は AI スキップ" if ev.get("ai_skip_now") else ""
                    status.write(f"`[{_ts()}]` ✖ {ev['name']} 失敗 ({ev['elapsed']:.1f}秒){note}")

            return status, sub_bar, _on_progress, _on_event, state

        if selections.get("scenario"):
            tick(f"テストシナリオ (AI, {len(ai_targets)}フォーム) を生成中…")
            status, sub_bar, _on_p, _on_e, _state = _make_ai_runner(
                f"テストシナリオ生成 ({len(ai_targets)}フォーム)"
            )
            try:
                client = _make_client()
                scenarios = generate_scenarios(study, ai_targets, client, _on_p, _on_e)
                wb = build_test_scenario_workbook(study, scenarios)
                files[f"{study.name}_テストシナリオ.xlsx"] = excel_to_bytes(wb)
                sub_bar.progress(1.0, text="完了")
                total_t = sum(_state["elapsed_per_sheet"])
                status.update(label=f"✅ テストシナリオ完了 (合計 {total_t:.1f}秒, {len(scenarios)}件)",
                              state="complete", expanded=False)
            except Exception as e:
                errors.append(("テストシナリオ", str(e)))
                status.update(label=f"❌ テストシナリオ失敗: {e}", state="error", expanded=True)

        if selections.get("manual"):
            tick(f"マニュアルチェックリスト (AI, {len(ai_targets)}フォーム) を生成中…")
            status, sub_bar, _on_p, _on_e, _state = _make_ai_runner(
                f"マニュアルチェック生成 ({len(ai_targets)}フォーム)"
            )
            try:
                client = _make_client()
                points = generate_check_points(study, ai_targets, client, _on_p, _on_e)
                wb = build_manual_check_workbook(study, points)
                files[f"{study.name}_マニュアルチェックリスト.xlsx"] = excel_to_bytes(wb)
                sub_bar.progress(1.0, text="完了")
                total_t = sum(_state["elapsed_per_sheet"])
                status.update(label=f"✅ マニュアルチェック完了 (合計 {total_t:.1f}秒, {len(points)}件)",
                              state="complete", expanded=False)
            except Exception as e:
                errors.append(("マニュアルチェックリスト", str(e)))
                status.update(label=f"❌ マニュアルチェック失敗: {e}", state="error", expanded=True)

        progress.progress(1.0, text="完了")

        if not files:
            st.warning("生成可能なドキュメントがありませんでした。")
            st.stop()

        # session_state に保持（DLボタン押下時の rerun でも消えないように）
        st.session_state["generated_files"] = files
        st.session_state["generated_study_name"] = study.name
        st.session_state["generated_ts"] = datetime.now().strftime("%Y%m%d_%H%M%S")
    except Exception as e:
        st.error(f"生成中にエラーが発生しました: {e}")
        st.exception(e)


# ---------------- 生成結果の表示（rerun されても残る） ----------------
if "generated_files" in st.session_state:
    files = st.session_state["generated_files"]
    gen_study_name = st.session_state["generated_study_name"]
    gen_ts = st.session_state["generated_ts"]

    st.success(f"✓ {len(files)} 件のドキュメントを生成しました。")

    head_c1, head_c2 = st.columns([4, 1])
    with head_c1:
        st.markdown(
            f'<div class="step-head" style="margin-top:14px;">'
            f'<div class="step-num" style="background:{GOLD};color:{NAVY};border-color:{NAVY};">↓</div>'
            f'<div class="step-title">生成結果</div></div>',
            unsafe_allow_html=True,
        )
    with head_c2:
        if st.button("結果をクリア", use_container_width=True, key="clear_results"):
            st.session_state.pop("generated_files", None)
            st.session_state.pop("generated_study_name", None)
            st.session_state.pop("generated_ts", None)
            st.rerun()

    for fname, fbytes in files.items():
        c1, c2 = st.columns([3, 1])
        with c1:
            st.markdown(
                f"""
<div class="result-file">
  <div class="ico">{_file_icon(fname)}</div>
  <div style="flex:1;min-width:0;">
    <div class="name" style="overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">{fname}</div>
    <div class="size">{_fmt_size(len(fbytes))}</div>
  </div>
</div>
""",
                unsafe_allow_html=True,
            )
        with c2:
            st.download_button(
                label="ダウンロード",
                data=fbytes,
                file_name=fname,
                mime=_mime_for(fname),
                use_container_width=True,
                key=f"dl_{fname}",
            )

    zip_bytes = files_to_zip_bytes(files)
    st.markdown("<div style='margin-top:10px;'></div>", unsafe_allow_html=True)
    st.download_button(
        label="📦 すべてを ZIP でダウンロード",
        data=zip_bytes,
        file_name=f"{gen_study_name}_EDC_docs_{gen_ts}.zip",
        mime="application/zip",
        use_container_width=True,
        type="primary",
        key="dl_zip",
    )
