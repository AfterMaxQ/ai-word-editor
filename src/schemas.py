# src/schemas.py

from typing import Any, Dict, List, Literal, Optional, Union, Annotated
from pydantic import BaseModel, Field, model_validator

# ==============================================================================
# SECTION 1: FINAL DOCUMENT STRUCTURE SCHEMA (No changes)
# ==============================================================================
class BaseProperties(BaseModel):
    class Config: extra = 'allow'
class ParagraphProperties(BaseProperties):
    style: Optional[str] = None; bold: Optional[bool] = None; bookmark_id: Optional[str] = None; alignment: Optional[Literal['left', 'center', 'right']] = None; font_name: Optional[str] = None; font_size: Optional[float] = None; first_line_indent: Optional[float] = None; font_color: Optional[str] = None; spacing_before: Optional[float] = None; spacing_after: Optional[float] = None; line_spacing: Optional[float] = None
class TableProperties(BaseProperties):
    header: Optional[bool] = None; alignments: Optional[List[str]] = None; style: Optional[str] = None; alignment: Optional[Literal['left', 'center', 'right']] = None
class HeaderFooterProperties(BaseProperties):
    text: str; alignment: Optional[Literal['left', 'center', 'right']] = 'center'
class TextRun(BaseModel): type: Literal['text']; text: str
class FootnoteRun(BaseModel): type: Literal['footnote']; text: str
class EndnoteRun(BaseModel): type: Literal['endnote']; text: str
class CrossReferenceRun(BaseModel): type: Literal['cross_reference']; target_bookmark: str
class FormulaRun(BaseModel): type: Literal['formula']; text: str # The LaTeX text
AnyRun = Annotated[Union[TextRun, FootnoteRun, EndnoteRun, CrossReferenceRun, FormulaRun], Field(discriminator='type')]

class ParagraphElement(BaseModel):
    type: Literal['paragraph']
    text: Optional[str] = None
    content: Optional[List[AnyRun]] = None
    properties: Optional[ParagraphProperties] = None

    @model_validator(mode='after')
    def check_text_or_content(self) -> 'ParagraphElement':
        if self.text is not None and self.content is not None:
            raise ValueError("Paragraph cannot have both 'text' and 'content'.")
        return self


class NumberingLevel(BaseModel):
    level: int = Field(..., description="The indentation level, starting from 0.")
    number_format: Literal['decimal', 'lowerLetter', 'upperLetter', 'lowerRoman', 'upperRoman', 'chineseCounting'] = Field(..., description="The format of the number itself.")
    text_format: str = Field(..., description="The text format string, e.g., '%1.' for level 0, '(%2)' for level 1.")

class NumberingDefinition(BaseModel):
    name: str = Field(..., description="A unique name for this numbering scheme, e.g., 'ThesisHeadings'.")
    style_links: Dict[str, int] = Field(..., description="A mapping of Style Name to a level index, e.g., {'Heading 1': 0, 'Heading 2': 1}.")
    levels: List[NumberingLevel]
class ListElement(BaseModel): type: Literal['list']; items: Optional[List[str]] = None; properties: Optional[Dict[str, Any]] = None
class TableElement(BaseModel): type: Literal['table']; data: Optional[List[List[str]]] = None; properties: Optional[TableProperties] = None
class ImageElement(BaseModel): type: Literal['image']; properties: Dict[str, Any]
class FormulaElement(BaseModel): type: Literal['formula']; properties: Dict[str, Any]
class HeaderElement(BaseModel): type: Literal['header']; properties: HeaderFooterProperties
class FooterElement(BaseModel): type: Literal['footer']; properties: HeaderFooterProperties
class PageBreakElement(BaseModel): type: Literal['page_break']
class ColumnBreakElement(BaseModel): type: Literal['column_break']
class TocElement(BaseModel): type: Literal['toc']; properties: Optional[Dict[str, Any]] = None
AnyElement = Annotated[Union[ParagraphElement, ListElement, TableElement, ImageElement, FormulaElement, HeaderElement, FooterElement, PageBreakElement, ColumnBreakElement, TocElement], Field(discriminator='type')]
class PageSetup(BaseModel):
    orientation: Optional[Literal['portrait', 'landscape']] = None; margins: Optional[Dict[str, float]] = None; endnote_number_format: Optional[str] = None; footnote_number_format: Optional[str] = None; endnote_reference_format: Optional[str] = None; footnote_reference_format: Optional[str] = None
class Section(BaseModel): properties: Optional[Dict[str, Any]] = None; elements: List[AnyElement] = Field(default_factory=list)
class NumberingLevel(BaseModel): level: int; number_format: Literal['decimal', 'lowerLetter', 'upperLetter', 'lowerRoman', 'upperRoman']; text_format: str
class NumberingDefinition(BaseModel): name: str; style_links: Dict[str, int]; levels: List[NumberingLevel]
class DocumentModel(BaseModel):
    page_setup: Optional[PageSetup] = None
    numbering_definitions: Optional[List[NumberingDefinition]] = None # <-- ADDED THIS LINE
    sections: Optional[List[Section]] = None

# ==============================================================================
# SECTION 2: LOGICAL COMMAND BLOCK SCHEMA (Phase 1 Agent Output)
# ==============================================================================
class LogicalCommandBlock(BaseModel):
    """
    【已更新 v2】定义了指令规整代理的输出结构。
    此版本使用唯一的字符串ID和基于ID的依赖关系，以降低AI的认知负荷并消除索引错误。
    """
    id: str = Field(..., description="一个为此块分配的简短、唯一、驼峰式的字符串ID，例如 'title' 或 'execSummary'。")
    primary_command: str = Field(..., description="核心的创建或动作型指令。")
    follow_up_commands: List[str] = Field(default_factory=list, description="用于修饰主指令的后续指令列表。")
    dependencies: Optional[List[str]] = Field(
        default_factory=list,
        description="一个字符串列表，表示此块所依赖的其他块的ID。"
    )

class CommandBlockContainer(BaseModel):
    """
    用于验证指令规整代理完整JSON输出的根模型。
    """
    command_blocks: List[LogicalCommandBlock]

# ==============================================================================
# SECTION 3: TOOL CALLING SCHEMA (Phase 2 Agent Output)
# ==============================================================================
class ToolCall(BaseModel):
    """
    【已重构】定义了工具调用的Schema。
    tool_name 严格与 DocumentBuilder 的方法名对齐。
    """
    tool_name: Literal[
        'add_paragraph', 'add_list', 'add_table', 'update_properties',
        'set_page_orientation', 'set_margins_cm', 'define_numbering',
        'add_header', 'add_footer', 'add_page_break', 'add_toc', 'no_op'
    ]
    tool_input: Dict[str, Any]

class ToolCallContainer(BaseModel):
    """
    用于验证工具调用代理完整JSON输出的根模型。
    """
    calls: List[ToolCall]