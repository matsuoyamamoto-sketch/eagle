"""アプリ設定。

ローカル: .env から読み込み。
Streamlit Community Cloud: st.secrets から読み込み (secrets.toml に値を設定)。
"""
from __future__ import annotations

import os
from pathlib import Path

from pydantic_settings import BaseSettings, SettingsConfigDict


def _load_streamlit_secrets() -> None:
    """Streamlit Cloud 環境では st.secrets を環境変数に転写する。"""
    try:
        import streamlit as st  # type: ignore
        if hasattr(st, "secrets"):
            for k, v in dict(st.secrets).items():
                os.environ.setdefault(k, str(v))
    except Exception:
        # streamlit が無い (CLI/test) 環境では何もしない
        pass


_load_streamlit_secrets()


class Settings(BaseSettings):
    cohere_api_key: str = ""
    cohere_model: str = "command-r-plus-08-2024"
    cohere_requests_per_minute: int = 15
    cohere_max_retries: int = 5

    output_dir: Path = Path("./output")
    log_level: str = "INFO"

    model_config = SettingsConfigDict(env_file=".env", env_file_encoding="utf-8", extra="ignore")


settings = Settings()
