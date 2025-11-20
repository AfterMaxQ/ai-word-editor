from typing import List, Dict, Any, Optional, Literal


class ParagraphProxy:
    """一个代理对象，用于对最新创建的段落进行链式操作。"""

    def __init__(self, paragraph_dict: Dict[str, Any]):
        self._paragraph = paragraph_dict
        if 'properties' not in self._paragraph:
            self._paragraph['properties'] = {}

    def set_alignment(self, alignment: Literal['left', 'center', 'right']) -> 'ParagraphProxy':
        """设置段落的对齐方式。"""
        self._paragraph['properties']['alignment'] = alignment
        return self

    def bookmark(self, bookmark_id: str) -> 'ParagraphProxy':
        """【名称变更】为段落添加书签。"""
        self._paragraph['properties']['bookmark_id'] = bookmark_id
        return self


class TableProxy:
    """一个代理对象，用于对最新创建的表格进行链式操作。"""

    def __init__(self, table_dict: Dict[str, Any]):
        self._table = table_dict
        if 'properties' not in self._table:
            self._table['properties'] = {}

    def set_alignment(self, alignment: Literal['left', 'center', 'right']) -> 'TableProxy':
        """设置表格的整体对齐方式。"""
        self._table['properties']['alignment'] = alignment
        return self


class DocumentBuilder:
    """
    一个领域特定语言（DSL）构建器，为AI提供一个简单的API来创建文档结构。
    它将所有操作封装为方法，内部维护一个符合最终Schema的JSON状态。
    """

    def __init__(self):
        self.doc_state: Dict[str, Any] = {
            "page_setup": {},
            "numbering_definitions": [],
            "sections": [{"elements": []}]
        }

    def _get_current_elements(self) -> List[Dict[str, Any]]:
        return self.doc_state["sections"][-1]["elements"]

    def _add_element(self, element: Dict[str, Any]):
        self._get_current_elements().append(element)

    def set_page_orientation(self, orientation: Literal['portrait', 'landscape']) -> 'DocumentBuilder':
        """设置页面方向。"""
        self.doc_state['page_setup']['orientation'] = orientation
        return self

    def set_margins_cm(self, top: float, bottom: float, left: float, right: float) -> 'DocumentBuilder':
        """以厘米为单位，设置页面边距。"""
        self.doc_state['page_setup']['margins'] = {
            "top": top, "bottom": bottom, "left": left, "right": right
        }
        return self

    def add_header(self, text: str, alignment: Literal['left', 'center', 'right'] = 'center') -> 'DocumentBuilder':
        """添加页眉。"""
        self._add_element({"type": "header", "properties": {"text": text, "alignment": alignment}})
        return self

    def add_footer(self, text: str, alignment: Literal['left', 'center', 'right'] = 'center') -> 'DocumentBuilder':
        """添加页脚。"""
        self._add_element({"type": "footer", "properties": {"text": text, "alignment": alignment}})
        return self

    def add_paragraph(self, text: str = "", style: Optional[str] = None) -> ParagraphProxy:
        """添加一个新段落，并返回一个可链式操作的代理对象。"""
        element = {"type": "paragraph", "text": text, "properties": {}}
        if style:
            element['properties']['style'] = style
        self._add_element(element)
        return ParagraphProxy(element)

    def add_list(self, items: List[str], ordered: bool = False) -> 'DocumentBuilder':
        """添加一个有序或无序列表。"""
        element = {"type": "list", "items": items, "properties": {"ordered": ordered}}
        self._add_element(element)
        return self

    def add_table(self, data: List[List[str]], style: Optional[str] = None) -> TableProxy:
        """添加一个新表格，并返回一个可链式操作的代理对象。"""
        element = {"type": "table", "data": data, "properties": {}}
        if style:
            element['properties']['style'] = style
        self._add_element(element)
        return TableProxy(element)

    def define_numbering(self, name: str, style_links: Dict[str, int],
                         levels: List[Dict[str, Any]]) -> 'DocumentBuilder':
        if 'numbering_definitions' not in self.doc_state or self.doc_state['numbering_definitions'] is None:
            self.doc_state['numbering_definitions'] = []

        definition = {
            "name": name,
            "style_links": style_links,
            "levels": levels
        }
        self.doc_state['numbering_definitions'].append(definition)
        return self

    def add_page_break(self) -> 'DocumentBuilder':
        """插入一个分页符。"""
        self._add_element({"type": "page_break"})
        return self

    def add_toc(self) -> 'DocumentBuilder':
        """插入一个目录。"""
        self._add_element({"type": "toc"})
        return self

    def get_document_state(self) -> Dict[str, Any]:
        """获取最终构建的文档状态JSON。"""
        return self.doc_state

    def get_element_by_bookmark(self, bookmark_id: str) -> Optional[Dict[str, Any]]:
        """
        Finds an element within the document state by its assigned bookmark ID.

        Args:
            bookmark_id (str): The unique ID of the bookmark to find.

        Returns:
            Optional[Dict[str, Any]]: The element dictionary if found, otherwise None.
        """
        for section in self.doc_state.get("sections", []):
            for element in section.get("elements", []):
                if element.get("properties", {}).get("bookmark_id") == bookmark_id:
                    return element
        return None

    def update_table(self, target_bookmark_id: str, new_data: List[List[str]],
                     action: Literal['append_rows', 'overwrite'] = 'overwrite') -> 'DocumentBuilder':
        """
        Updates an existing table identified by its bookmark.

        Args:
            target_bookmark_id (str): The bookmark ID of the table to modify.
            new_data (List[List[str]]): The new data for the table.
            action (Literal['append_rows', 'overwrite']): The action to perform.
                'overwrite' replaces the entire table data.
                'append_rows' adds the new_data as new rows.

        Returns:
            DocumentBuilder: The builder instance for chaining.
        """
        table_element = self.get_element_by_bookmark(target_bookmark_id)
        if table_element and table_element.get("type") == "table":
            if action == 'overwrite':
                table_element['data'] = new_data
            elif action == 'append_rows':
                if 'data' not in table_element or table_element['data'] is None:
                    table_element['data'] = []
                table_element['data'].extend(new_data)
        else:
            # Add a warning paragraph if the table is not found
            self._add_element({
                "type": "paragraph",
                "text": f"[AI Error: Could not find table with bookmark '{target_bookmark_id}' to update.]"
            })
        return self