"""EDC設定 JSON の Pydantic モデル。

実データ (Bev-FOLFOX-SBC) の構造調査に基づき、必要最小限のフィールドのみ厳密に型付けし、
未知フィールドは ``extra="allow"`` で取りこぼさない方針。
"""
from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, ConfigDict, Field

FieldItemType = Literal[
    "FieldItem::Heading",
    "FieldItem::Article",
    "FieldItem::Assigned",
    "FieldItem::Reference",
    "FieldItem::Note",
    "FieldItem::Allocation",
]


class _Base(BaseModel):
    model_config = ConfigDict(extra="allow", populate_by_name=True)


# ---------- validators ----------
class PresenceValidator(_Base):
    validate_presence_id: str | None = None


class DateValidator(_Base):
    validate_date_after_or_equal_to: str | None = None
    validate_date_before_or_equal_to: str | None = None
    validate_date_after: str | None = None
    validate_date_before: str | None = None


class NumericalityValidator(_Base):
    validate_numericality_less_than_or_equal_to: str | None = None
    validate_numericality_greater_than_or_equal_to: str | None = None
    validate_numericality_less_than: str | None = None
    validate_numericality_greater_than: str | None = None
    validate_numericality_equal_to: str | None = None


class FormulaValidator(_Base):
    """エディットチェックの本体。`if` が条件式、`message` がエラーメッセージ。"""

    validate_formula_if: str | None = None
    validate_formula_message: str | None = None


class Validators(_Base):
    presence: PresenceValidator | None = None
    date: DateValidator | None = None
    numericality: NumericalityValidator | None = None
    formula: FormulaValidator | None = None

    def is_empty(self) -> bool:
        return not any([self.presence, self.date, self.numericality, self.formula])


# ---------- field item ----------
class FieldItem(_Base):
    name: str  # field1, field2, ...
    label: str = ""
    description: str = ""
    seq: int = 0
    type: FieldItemType
    is_invisible: bool = False
    field_type: str | None = None  # date / text / radio_button / drug / meddra / sae_report
    default_value: str | None = None
    validators: Validators = Field(default_factory=Validators)
    option_name: str | None = None         # FieldItem::Assigned のコードリスト参照
    reference_field: str | None = None     # FieldItem::Reference の "sheet.field"
    reference_type: str | None = None
    formula_field: str | None = None
    content: str | None = None             # FieldItem::Note 本文
    level: int | None = None
    deviation: str | None = None
    link_type: str | None = None
    normal_range: dict[str, Any] = Field(default_factory=dict)


class CdiscSheetConfig(_Base):
    prefix: str | None = None
    label: str | None = None
    table: dict[str, str | None] = Field(default_factory=dict)


class Sheet(_Base):
    name: str
    alias_name: str | None = None
    category: str | None = None
    is_serious: bool = False
    is_closed: bool = False
    field_items: list[FieldItem] = Field(default_factory=list)
    cdisc_sheet_configs: list[CdiscSheetConfig] = Field(default_factory=list)


# ---------- options (codelist) ----------
class OptionValue(_Base):
    name: str
    seq: int = 0
    code: str = ""
    is_usable: bool = True


class CodeList(_Base):
    uuid: str | None = None
    name: str
    is_extensible: bool = False
    values: list[OptionValue] = Field(default_factory=list)


# ---------- sheet group ----------
class SheetGroup(_Base):
    uuid: str | None = None
    name: str
    alias_name: str | None = None
    allocation_group: str | None = None
    is_default: bool = False
    sheets: list[dict[str, Any]] = Field(default_factory=list)


# ---------- top level ----------
class Study(_Base):
    name: str
    proper_name: str = ""
    disease_category: str | None = None
    sdtm_version: str | None = None
    sdtm_terminology_version: str | None = None
    ctcae_version: str | None = None
    organization_name: str | None = None
    is_observation_study: bool = False
    is_pv_enabled: bool = False
    is_registrational: bool = False
    uuid: str | None = None

    sheet_groups: list[SheetGroup] = Field(default_factory=list)
    sheets: list[Sheet] = Field(default_factory=list)
    options: list[CodeList] = Field(default_factory=list)
    visits: list[dict[str, Any]] = Field(default_factory=list)
    visit_groups: list[dict[str, Any]] = Field(default_factory=list)
    epro_questionnaires: list[dict[str, Any]] = Field(default_factory=list)
    sheet_orders: list[dict[str, Any]] = Field(default_factory=list)

    # ---------- 集計ヘルパ ----------
    def total_field_items(self) -> int:
        return sum(len(s.field_items) for s in self.sheets)

    def count_validator(self, kind: Literal["presence", "date", "numericality", "formula"]) -> int:
        n = 0
        for s in self.sheets:
            for fi in s.field_items:
                if getattr(fi.validators, kind) is not None:
                    n += 1
        return n

    def count_reference_items(self) -> int:
        return sum(
            1 for s in self.sheets for fi in s.field_items if fi.type == "FieldItem::Reference"
        )

    def codelist_by_name(self, name: str) -> CodeList | None:
        for o in self.options:
            if o.name == name:
                return o
        return None
