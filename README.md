# EAGLE — EDC Auto Generator & Logic Extractor

EDC設定JSONから以下5種のドキュメントを自動生成する Streamlit アプリ。

1. EDC仕様書 (Excel)
2. バリデーションプラン (Word)
3. エディットチェック確認書 (Excel)
4. テストシナリオ (Excel, AI生成)
5. マニュアルチェックリスト (Excel, AI生成)

---

## ローカル起動

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
copy .env.example .env
# .env を編集して COHERE_API_KEY を設定 (UI のサイドバーから入力でも可)
streamlit run app.py
```

ブラウザで http://localhost:8501

---

## Streamlit Community Cloud デプロイ

### 0. 前提
- 無料枠は **Public リポジトリのみ**。秘匿コードがある場合は Streamlit Cloud Pro / 別ホスティングを検討。
- API キーは **コミットしない**。`st.secrets` に登録する。

### 1. GitHub に push
```bash
git init
git add .
git commit -m "initial: EAGLE"
git branch -M main
git remote add origin https://github.com/<your-account>/eagle.git
git push -u origin main
```

`.gitignore` で `.env` / `output/` / `*.xlsx` 等は除外済み。

### 2. Streamlit Cloud で New app
- https://share.streamlit.io/ にログイン
- "New app" → リポジトリ・ブランチ・`app.py` を選択
- "Deploy"

### 3. Secrets を設定
アプリ管理画面 → "Settings" → "Secrets" に以下を貼り付け（`.streamlit/secrets.toml.example` 参照）:

```toml
COHERE_API_KEY = ""             # 共有キー使う場合のみ。空なら各ユーザー入力
COHERE_MODEL = "command-r-plus-08-2024"
COHERE_REQUESTS_PER_MINUTE = 15
COHERE_MAX_RETRIES = 5
```

### 運用パターン

| パターン | Secrets の API Key | 利用者の操作 |
|---|---|---|
| **共有キー** | 設定する | そのまま使える |
| **各自キー** | 空欄のまま | サイドバーで自分のキーを入力 |

ユーザーが各自キーを入れる運用なら Secrets は空でも動作します（サイドバー入力が優先される）。

---

## ディレクトリ

```
eagle/
├── app.py                  # Streamlit エントリ
├── src/
│   ├── parser/             # JSON → Pydantic モデル
│   ├── generators/         # 各ドキュメント生成
│   ├── ai/                 # Cohere 連携
│   └── utils/
├── .streamlit/             # Streamlit 設定 (config.toml)
├── samples/
├── tests/
└── output/                 # 生成物 (gitignore)
```

---

## テスト

```bash
.venv\Scripts\activate
pytest tests/ -v
```

サンプル JSON が無い環境ではテストは自動 skip されます。
