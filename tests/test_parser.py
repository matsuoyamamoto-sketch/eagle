"""edc_parser の単体テスト。実JSON (Bev-FOLFOX-SBC) を投入して期待値と一致するか確認。"""
from __future__ import annotations

from pathlib import Path

import pytest

from src.parser.edc_parser import load_study
from src.parser.models import Study

SAMPLE_JSON = Path(
    r"c:/Users/matsu/Box/Datacenter/ISR/Ptosh/検証/JSON/Bev-FOLFOX-SBC_250929_1501.json"
)


@pytest.fixture(scope="module")
def study() -> Study:
    if not SAMPLE_JSON.exists():
        pytest.skip(f"sample JSON not found: {SAMPLE_JSON}")
    return load_study(SAMPLE_JSON)


def test_study_meta(study: Study):
    assert study.name == "Bev-FOLFOX-SBC"
    assert "小腸癌" in study.proper_name
    assert study.organization_name == "kanribev"
    assert study.sdtm_version == "3.2"


def test_counts(study: Study):
    assert len(study.sheets) == 162
    assert len(study.sheet_groups) == 3
    assert len(study.options) == 85
    assert study.total_field_items() == 32716


def test_validator_counts(study: Study):
    # JSON 構造調査時の実測値
    assert study.count_validator("presence") == 6809
    assert study.count_validator("date") == 3112
    assert study.count_validator("numericality") == 2817
    assert study.count_validator("formula") == 4065


def test_field_item_types(study: Study):
    types = {fi.type for s in study.sheets for fi in s.field_items}
    assert types == {
        "FieldItem::Heading",
        "FieldItem::Article",
        "FieldItem::Assigned",
        "FieldItem::Reference",
        "FieldItem::Note",
        "FieldItem::Allocation",
    }


def test_formula_extraction(study: Study):
    """formula validator が条件式とメッセージを保持していること。"""
    found = []
    for s in study.sheets:
        for fi in s.field_items:
            if fi.validators.formula and fi.validators.formula.validate_formula_if:
                found.append(fi.validators.formula)
    assert len(found) > 0
    # 最初に見つかる formula は同意取得日の年齢チェック
    assert any("age(" in (f.validate_formula_if or "") for f in found)


def test_codelist_lookup(study: Study):
    cl = study.codelist_by_name("Sex")
    assert cl is not None
    codes = {v.code for v in cl.values}
    assert "M" in codes or "MALE" in codes or len(codes) > 0
