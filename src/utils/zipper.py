"""生成物を ZIP に固める。"""
from __future__ import annotations

import io
import zipfile
from pathlib import Path


def files_to_zip_bytes(files: dict[str, bytes]) -> bytes:
    """{filename: bytes} を受け取り ZIP のバイト列を返す。"""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for name, data in files.items():
            zf.writestr(name, data)
    return buf.getvalue()


def excel_to_bytes(workbook) -> bytes:
    buf = io.BytesIO()
    workbook.save(buf)
    return buf.getvalue()


def docx_to_bytes(document) -> bytes:
    buf = io.BytesIO()
    document.save(buf)
    return buf.getvalue()
