"""EDC設定 JSON のロード&パース。"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any, BinaryIO, Union

from .models import Study


def load_study(source: Union[str, Path, BinaryIO, bytes, dict[str, Any]]) -> Study:
    """JSONファイルパス / バイト列 / file-like / dict から Study を生成。"""
    if isinstance(source, dict):
        data = source
    elif isinstance(source, (bytes, bytearray)):
        data = json.loads(source.decode("utf-8"))
    elif isinstance(source, (str, Path)):
        with open(source, "r", encoding="utf-8") as f:
            data = json.load(f)
    else:
        # file-like (Streamlit UploadedFile 等)
        raw = source.read()
        if isinstance(raw, bytes):
            raw = raw.decode("utf-8")
        data = json.loads(raw)
    return Study.model_validate(data)
