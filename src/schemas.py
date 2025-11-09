# src/schemas.py
from pydantic import BaseModel, Field, model_validator
from typing import List, Optional, Literal, Union, Dict, Any, Annotated # <-- Import Annotated

# --- 首先定义最内层的、可复用的模型 ---

class BaseProperties(BaseModel):
    """一个基础的属性模型，允许额外的属性"""
    class Config:
        extra = 'allow'

class ParagraphProperties(BaseProperties):
    style: Optional[str] = None
    font_name: Optional[str] = None
    font_size: Optional[int] = None
    bold: Optional[bool] = None
    first_line_indent: Optional[float] = None
    bookmark_id: Optional[str] = None
    font_color: Optional[str] = Field(None, description="字体的颜色，使用 RRGGBB 十六进制格式。")
    spacing_before: Optional[float] = Field(None, description="段落前的间距，单位是磅 (pt)。")
    spacing_after: Optional[float] = Field(None, description="段落后的间距，单位是磅 (pt)。")
    line_spacing: Optional[float] = Field(None, description="行距。例如 1.0 代表单倍行距, 1.5 代表1.5倍行距。")

# --- 定义 "Run" 对象 (段落内部的组件) ---

class TextRun(BaseModel):
    type: Literal['text']
    text: str

class FootnoteRun(BaseModel):
    type: Literal['footnote']
    text: str

class EndnoteRun(BaseModel):
    type: Literal['endnote']
    text: str

class CrossReferenceRun(BaseModel):
    type: Literal['cross_reference']
    target_bookmark: str

# ▼▼▼ [FIX] ▼▼▼
# The discriminator is now correctly applied to the Union type itself using Annotated.
AnyRun = Annotated[
    Union[TextRun, FootnoteRun, EndnoteRun, CrossReferenceRun],
    Field(discriminator='type')
]


# --- 定义顶层的 "Element" 对象 ---

class ParagraphElement(BaseModel):
    type: Literal['paragraph']
    text: Optional[str] = None
    # The list now contains the Annotated AnyRun type
    content: Optional[List[AnyRun]] = None
    properties: Optional[ParagraphProperties] = None

    @model_validator(mode='after')
    def check_text_or_content(self) -> 'ParagraphElement':
        if self.text is not None and self.content is not None:
            raise ValueError("A paragraph cannot have both 'text' and 'content'. Use 'content' for complex structures.")
        return self

class ListElement(BaseModel):
    type: Literal['list']
    items: List[str]
    properties: Optional[Dict[str, Any]] = None

class TableElement(BaseModel):
    type: Literal['table']
    data: List[List[str]]
    properties: Optional[Dict[str, Any]] = None

class ImageElement(BaseModel):
    type: Literal['image']
    properties: Dict[str, Any]

class FormulaElement(BaseModel):
    type: Literal['formula']
    properties: Dict[str, Any]

class HeaderElement(BaseModel):
    type: Literal['header']
    properties: Dict[str, Any]

class FooterElement(BaseModel):
    type: Literal['footer']
    properties: Dict[str, Any]

class PageBreakElement(BaseModel):
    type: Literal['page_break']

class ColumnBreakElement(BaseModel):
    type: Literal['column_break']

class TocElement(BaseModel):
    type: Literal['toc']
    properties: Optional[Dict[str, Any]] = None

# ▼▼▼ [FIX] ▼▼▼
# Same fix as AnyRun: apply the discriminator to the Union of elements.
AnyElement = Annotated[
    Union[
        ParagraphElement, ListElement, TableElement, ImageElement, FormulaElement,
        HeaderElement, FooterElement, PageBreakElement, ColumnBreakElement, TocElement
    ],
    Field(discriminator='type')
]


# --- 定义最高层的文档结构 ---

class PageSetup(BaseModel):
    orientation: Optional[str] = None
    margins: Optional[Dict[str, float]] = None
    endnote_number_format: Optional[str] = None
    footnote_number_format: Optional[str] = None
    endnote_reference_format: Optional[str] = None
    footnote_reference_format: Optional[str] = None

class Section(BaseModel):
    properties: Optional[Dict[str, Any]] = None
    # The list now correctly contains the Annotated AnyElement type
    elements: List[AnyElement]

class DocumentModel(BaseModel):
    """这是我们期望AI返回的完整JSON结构对应的Pydantic模型"""
    page_setup: Optional[PageSetup] = None
    sections: Optional[List[Section]] = None