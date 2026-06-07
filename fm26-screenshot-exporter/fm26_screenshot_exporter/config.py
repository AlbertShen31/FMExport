"""Pydantic models for extraction profiles and pipeline configuration."""

from __future__ import annotations

from enum import Enum
from typing import Literal

from pydantic import BaseModel, Field, field_validator


class ColumnStrategy(str, Enum):
    FROM_HEADER_OCR = "from_header_ocr"
    FIXED_BOUNDARIES = "fixed_boundaries"
    INFER_FROM_VERTICAL_ALIGNMENT = "infer_from_vertical_alignment"


class CropRegion(BaseModel):
    x: int = 0
    y: int = 0
    width: int
    height: int

    @field_validator("width", "height")
    @classmethod
    def positive_dimensions(cls, value: int) -> int:
        if value <= 0:
            raise ValueError("width and height must be positive")
        return value


class FixedColumn(BaseModel):
    name: str
    x_start: int
    x_end: int

    @field_validator("x_end")
    @classmethod
    def end_after_start(cls, value: int, info) -> int:
        x_start = info.data.get("x_start", 0)
        if value <= x_start:
            raise ValueError("x_end must be greater than x_start")
        return value


class HeaderYRange(BaseModel):
    y_start: int
    y_end: int


NormalizerType = Literal[
    "age",
    "number",
    "percentage",
    "currency",
    "wage",
    "date",
    "star_rating",
    "position",
    "nationality",
    "club",
    "name",
]


class Profile(BaseModel):
    name: str
    expected_columns: list[str] = Field(default_factory=list)
    manual_crop: CropRegion | None = None
    header_y_range: HeaderYRange | None = None
    row_height: int | None = None
    row_y_tolerance: int = 12
    min_confidence: float = 50.0
    column_strategy: ColumnStrategy = ColumnStrategy.FIXED_BOUNDARIES
    fixed_columns: list[FixedColumn] = Field(default_factory=list)
    value_normalizers: dict[str, NormalizerType] = Field(default_factory=dict)
    upscale_factor: float = 2.0
    required_columns: list[str] = Field(default_factory=list)
    min_rows: int = 1
    max_empty_cell_ratio: float = 0.5
    min_avg_confidence: float = 40.0

    @field_validator("min_confidence", "min_avg_confidence")
    @classmethod
    def confidence_range(cls, value: float) -> float:
        if not 0 <= value <= 100:
            raise ValueError("confidence must be between 0 and 100")
        return value


class OcrToken(BaseModel):
    text: str
    confidence: float
    x: int
    y: int
    width: int
    height: int
    line_num: int
    block_num: int
    par_num: int

    @property
    def x_center(self) -> float:
        return self.x + self.width / 2

    @property
    def y_center(self) -> float:
        return self.y + self.height / 2


class TableRegions(BaseModel):
    table_bbox: tuple[int, int, int, int]
    header_bbox: tuple[int, int, int, int] | None = None
    row_y_boundaries: list[int] = Field(default_factory=list)
    column_x_boundaries: list[int] = Field(default_factory=list)


class ParsedRow(BaseModel):
    row_index: int
    cells: dict[str, str]
    normalized_cells: dict[str, object | None] = Field(default_factory=dict)
    avg_confidence: float | None = None
    low_confidence: bool = False


class ParseResult(BaseModel):
    source_path: str
    profile_name: str
    columns: list[str]
    rows: list[ParsedRow]
    ocr_tokens: list[OcrToken]
    table_regions: TableRegions
    metadata: dict[str, object] = Field(default_factory=dict)
