# src/doc_generater.py

def load_document_data(filepath):
    """
    从指定的JSON文件中加载并解析文档结构数据。

    Args:
        filepath (str): 包含文档结构定义的JSON文件的路径。

    Returns:
        dict: 一个包含文档完整结构数据的字典。如果文件未找到或JSON格式无效，函数会打印错误信息并终止程序。
    """

def apply_paragraph_properties(paragraph, properties: dict):
    """
    将一个属性字典中定义的格式应用到给定的 `python-docx` 段落对象上。

    Args:
        paragraph: 一个 `python-docx` 的 Paragraph 对象。
        properties (dict): 一个包含段落格式化指令的字典。

    数据处理过程:
        1. 获取段落的 `paragraph_format` 对象以应用段落级别的属性。
        2. 检查并应用 `spacing_before` (段前距), `spacing_after` (段后距), `line_spacing` (行距) 和 `first_line_indent` (首行缩进)。所有单位（如磅、厘米）都会被转换为`python-docx`内部所需的格式。
        3. 遍历段落中的每一个 `run` 对象（文本块）。
        4. 对每个 `run` 的 `font` 对象应用字符级别的属性，包括 `font_name` (字体名称，同时设置东亚字体), `font_size` (字号), `bold` (加粗), 和 `font_color` (字体颜色，从十六进制字符串转换)。
        5. 包含对无效颜色代码的警告处理。
    """

def add_table_from_data(doc, element: dict):
    """
    根据字典数据在Word文档中创建一个完整的表格。

    Args:
        doc: The python-docx Document object.
        element (dict): 一个描述表格的字典，必须包含 `data` 键，可选 `properties` 键。

    数据处理过程:
        1. 从 `element` 字典中提取表格数据 (`data`) 和属性 (`properties`)。
        2. 验证 `data` 是否有效，如果为空则跳过。
        3. 根据数据的第一行推断表格的列数，并创建表格。
        4. 如果 `properties` 中 `header` 为 `true`，则将第一行作为表头填充，并自动加粗。
        5. 遍历剩余的数据行，动态地向表格中添加新行并填充单元格内容。
        6. 根据 `properties` 中的 `alignments` 数组，为表格的每一列设置文本对齐方式（左、中、右）。
    """

def add_list_from_data(doc, element: dict):
    """
    根据字典数据在Word文档中添加一个有序列表或无序列表。

    Args:
        doc: The python-docx Document object.
        element (dict): 描述列表的字典，包含 `items` 数组和可选的 `properties`。

    数据处理过程:
        1. 从 `element` 字典中提取列表项 (`items`) 和属性 (`properties`)。
        2. 检查 `properties` 中的 `ordered` 键。如果为 `true`，则使用 'List Number' 样式；否则使用 'List Bullet' 样式。
        3. 遍历 `items` 数组中的每个字符串，并使用选定的样式将其作为新段落添加到文档中。
    """

def add_image_from_data(doc, element: dict):
    """
    根据字典数据在Word文档中插入一张图片。

    Args:
        doc: The python-docx Document object.
        element (dict): 描述图片的字典，其 `properties` 必须包含 `path`，可选 `width` 和 `height`。

    数据处理过程:
        1. 从 `element['properties']` 中提取图片路径、宽度和高度。
        2. 验证路径是否存在。
        3. 将厘米单位的宽度和高度转换为 `python-docx` 的 `Cm` 对象。
        4. 使用 `doc.add_picture()` 方法插入图片，并应用指定的尺寸。
        5. 包含文件未找到和其他插入异常的错误处理。
    """

def add_page_number(paragraph):
    """
    在一个段落中通过直接操作底层OXML来插入一个动态的页码域。

    Args:
        paragraph: The python-docx Paragraph object to which the page number will be added.

    数据处理过程:
        这是一个低级函数，它不依赖于简单的API调用。
        1. 在段落中创建一个新的 run。
        2. 创建三个OXML元素：
           - `<w:fldChar w:fldCharType="begin">`: 标记一个复杂域的开始。
           - `<w:instrText xml:space="preserve">PAGE</w:instrText>`: 这是Word的指令文本，`PAGE` 表示当前页码。
           - `<w:fldChar w:fldCharType="end">`: 标记域的结束。
        3. 将这三个元素按顺序附加到新创建的 run 的XML表示 (`_r`) 中。
    """

def add_header_from_data(doc, element: dict, section):
    """
    为一个特定的文档节(section)添加页眉。

    Args:
        doc: The python-docx Document object.
        element (dict): 描述页眉的字典，包含 `properties` (text, alignment)。
        section: The python-docx Section object to which the header will be applied.

    数据处理过程:
        1. 获取指定节的 `header` 对象。
        2. 清空页眉中已有的内容，以确保是全新的。
        3. 根据 `properties` 设置段落的对齐方式。
        4. 检查页眉文本中是否包含 `{PAGE_NUM}` 占位符。
        5. 如果包含，则将文本分割成三部分（占位符前、页码、占位符后），并调用 `add_page_number()` 在中间插入动态页码。
        6. 如果不包含，则直接添加纯文本。
    """

def add_footer_from_data(doc, element: dict, section):
    """
    为一个特定的文档节(section)添加页脚。功能和逻辑与 `add_header_from_data` 完全相同，只是操作对象是节的 `footer`。

    Args:
        doc: The python-docx Document object.
        element (dict): 描述页脚的字典，包含 `properties` (text, alignment)。
        section: The python-docx Section object to which the footer will be applied.

    Returns:
        无。
    """

def post_process_footnotes(docx_bytes, footnotes_to_inject, reference_format: str = "#"):
    """
    **[核心函数]** 通过直接操作DOCX文件内部的XML，为文档添加脚注。`python-docx`本身不支持创建脚注，此函数是必需的后处理步骤。

    Args:
        docx_bytes (bytes): 包含脚注占位符（如 `__FOOTNOTE_...__`）的原始DOCX文件的字节流。
        footnotes_to_inject (dict): 一个字典，键是占位符，值是对应的脚注文本。
        reference_format (str): 脚注引用标记的格式，'#' 是数字的占位符 (例如 "[#]")。

    Returns:
        bytes: 经过修改后，包含真实脚注的新的DOCX文件的字节流。

    数据处理过程:
        这是一个复杂的底层操作，模拟了Word处理脚注的过程：
        1.  将输入的 `docx_bytes` 解压到一个临时目录中。
        2.  定位并解析几个关键的XML文件：`word/document.xml` (主内容), `word/footnotes.xml` (脚注内容，如果不存在则创建), `word/_rels/document.xml.rels` (关系定义), `[Content_Types].xml` (文件类型定义)。
        3.  遍历 `footnotes_to_inject` 字典中的每一个待注入的脚注。
        4.  **对于每个脚注**：
            a. 在 `footnotes.xml` 中创建一个新的 `<w:footnote>` XML元素，为其分配一个唯一的ID。元素内部包含标准的Word脚注结构，并将脚注文本填充进去。
            b. 根据 `reference_format` 参数（如 `[#]`），构建一个复杂的XML片段。这个片段包含上标格式，并将实际的脚注引用 `<w:footnoteReference>` 插入到 `#` 的位置。
            c. 在 `document.xml` 的内容中，搜索对应的占位符字符串，并将其替换为上一步生成的XML引用片段。
        5.  **更新元数据**：
            a. 在 `document.xml.rels` 中添加一个关系条目，将主文档与 `footnotes.xml` 文件关联起来。
            b. 在 `[Content_Types].xml` 中声明 `footnotes.xml` 的内容类型。
        6.  将临时目录中的所有文件重新打包成一个新的ZIP文件（即DOCX文件），并将其作为字节流返回。
    """

def post_process_endnotes(docx_bytes, endnotes_to_inject, reference_format: str = "#"):
    """
    **[核心函数]** 通过直接操作DOCX文件内部的XML，为文档添加尾注。此函数与 `post_process_footnotes` 的逻辑和处理流程几乎完全相同，但操作的是与尾注相关的文件和XML标签（如 `endnotes.xml`, `<w:endnote>`, `<w:endnoteReference>`）。

    Args:
        docx_bytes (bytes): 包含尾注占位符的原始DOCX文件的字节流。
        endnotes_to_inject (dict): 一个字典，键是占位符，值是对应的尾注文本。
        reference_format (str): 尾注引用标记的格式，'#' 是数字的占位符。

    Returns:
        bytes: 经过修改后，包含真实尾注的新的DOCX文件的字节流。
    """

def add_bookmark(paragraph, bookmark_name: str):
    """
    通过直接操作OXML，为一个段落的完整内容添加书签。

    Args:
        paragraph: The python-docx Paragraph object.
        bookmark_name (str): 书签的唯一名称。

    数据处理过程:
        1. 生成一个全局唯一的数字ID。
        2. 创建一个 `<w:bookmarkStart>` OXML元素，并设置其 `w:id` 和 `w:name` 属性。
        3. 将此起始标签插入到段落第一个 run 的XML (`_r`) 之前。
        4. 创建一个 `<w:bookmarkEnd>` OXML元素，并设置其 `w:id`。
        5. 将此结束标签插入到段落最后一个 run 的XML (`_r`) 之后。
    """

def add_cross_reference_field(paragraph, bookmark_name: str, display_text: str):
    """
    在段落中插入一个指向书签的交叉引用域（REF field），并提供一个预填充的显示文本。

    Args:
        paragraph: The python-docx Paragraph object.
        bookmark_name (str): 目标书签的名称。
        display_text (str): 在Word中更新域之前显示的缓存文本。

    数据处理过程:
        此函数通过创建一系列特定的 run 和 OXML 元素来构建一个Word可以识别的复杂域：
        1.  添加一个 `run`，其中包含 `<w:fldChar w:fldCharType="begin">` 来标记域的开始。
        2.  添加一个 `run`，其中包含指令文本 `<w:instrText> REF {bookmark_name} \h </w:instrText>`。
        3.  添加一个 `run`，其中包含 `<w:fldChar w:fldCharType="separate">` 来分隔指令和结果。
        4.  添加一个包含 `display_text` 的普通 `run`，作为域的缓存结果。
        5.  添加一个 `run`，其中包含 `<w:fldChar w:fldCharType="end">` 来标记域的结束。
    """

def apply_numbering_formats(docx_bytes, page_setup_data):
    """
    **[核心后处理函数]** 通过修改 `word/settings.xml` 文件来应用自定义的脚注和尾注编号格式（如罗马数字、字母等）。

    Args:
        docx_bytes (bytes): 原始DOCX文件的字节流。
        page_setup_data (dict): 包含 `endnote_number_format` 和 `footnote_number_format` 键的页面设置字典。

    Returns:
        bytes: 经过修改后，应用了新编号格式的DOCX文件的字节流。

    数据处理过程:
        1.  将输入的 `docx_bytes` 解压到临时目录。
        2.  定位并解析 `word/settings.xml` 文件。
        3.  在XML树中查找 `<w:endnotePr>` 和 `<w:footnotePr>` 元素（如果不存在则创建）。
        4.  在这些元素内部，查找或创建 `<w:numFmt>` 元素，并将其 `w:val` 属性设置为字典中指定的格式字符串（如 "upperRoman"）。
        5.  将修改后的 `settings.xml` 写回。
        6.  将临时目录重新打包成字节流并返回。
    """

def apply_page_setup(doc, page_setup_data: dict):
    """
    对文档的第一个节应用全局页面设置。

    Args:
        doc: The python-docx Document object.
        page_setup_data (dict): 包含页面设置（方向、边距）的字典。

    数据处理过程:
        1. 获取文档的第一个 `section` 对象。
        2. 根据 `orientation` 的值（如 'landscape'）设置页面方向，并手动交换页面宽高以确保生效。
        3. 根据 `margins` 字典中的 `top`, `bottom`, `left`, `right` 值，设置页边距。
    """

def apply_section_properties(section, section_data: dict):
    """
    对一个特定的文档节应用属性，主要是分栏设置。

    Args:
        section: The python-docx Section object.
        section_data (dict): 包含该节属性的字典，如 `properties: {"columns": 2}`。

    数据处理过程:
        这是一个低级OXML操作：
        1. 获取节的属性XML元素 `_sectPr`。
        2. 如果 `properties` 中定义了 `columns` 且大于1，则创建一个 `<w:cols>` OXML元素。
        3. 设置 `<w:cols>` 元素的 `w:num` 属性为指定的栏数。
        4. 将此 `<w:cols>` 元素附加到 `_sectPr` 中。
    """

def add_page_break_from_data(doc, element: dict):
    """
    在文档中添加一个分页符。

    Args:
        doc: The python-docx Document object.
        element (dict): 一个空的占位符字典。
    """

def add_column_break_from_data(doc, element: dict):
    """
    在文档中添加一个分栏符。

    Args:
        doc: The python-docx Document object.
        element (dict): 一个空的占位符字典。
    """

def add_toc_from_data(doc, element: dict):
    """
    通过直接操作OXML，在文档中插入一个目录（TOC）域。

    Args:
        doc: The python-docx Document object.
        element (dict): 描述目录的字典，可选 `properties.title`。

    数据处理过程:
        1. 如果提供了标题，则先添加一个标题段落。
        2. 创建一个新段落，并通过构建一系列OXML元素（`fldChar` begin, `instrText` with 'TOC ...' command, `fldChar` separate, `fldChar` end）来插入一个复杂的TOC域。
    """

def get_formula_xml_and_placeholder(element: dict) -> tuple[str | None, str | None]:
    """
    **[核心函数]** 将LaTeX公式字符串转换为OMML（Office Math Markup Language）XML。它采用了一个健壮的、多层次的转换策略。

    Args:
        element (dict): 描述公式的字典，`properties.text` 中包含LaTeX字符串。

    Returns:
        tuple[str | None, str | None]: 返回一个元组，第一个元素是用于文档的唯一占位符字符串，第二个元素是转换后的OMML XML字符串。如果转换完全失败，第二个元素可能是包含错误信息的有效XML。

    数据处理过程（多层防御策略）:
        1.  **尝试本地转换器**: 首先调用 `latex_to_omml` (一个基于规则的解析器) 进行转换。
        2.  **零信任验证**: 如果本地转换器返回了结果，会尝试用 `lxml` 重新解析该结果，以确保其XML语法有效。
        3.  **回退到LLM**: 如果本地转换失败或验证失败，则调用 `translate_latex_to_omml_llm` 函数，将转换任务交给大语言模型。
        4.  **再次验证**: 对LLM返回的XML也进行零信任验证，确保其是有效的XML且结构符合规范。
        5.  **最终错误处理**: 如果所有方法都失败，它会生成一个包含红色错误提示文本的、但本身是有效XML的片段。这可以防止因公式转换失败而导致整个Word文档损坏。
        6.  生成一个唯一的占位符字符串（如 `__FORMULA_...__`）与最终的XML一起返回。
    """

def create_document(data: dict) -> tuple[bytes | None, str | None]:
    """
    **[核心函数 - 总指挥]** 根据完整的、经过验证的JSON结构数据，生成一个功能齐全的Word文档。

    Args:
        data (dict): 描述整个文档的JSON结构。

    Returns:
        tuple[bytes | None, str | None]: 返回一个元组，第一个元素是最终生成的 `.docx` 文件的字节流，第二个元素是用于调试的、文档主体的最终XML内容的字符串。

    数据处理过程（一个复杂的多阶段流程）:
        1.  **初始化**: 创建一个空白的 `python-docx` Document对象，并初始化用于存储公式、脚注、尾注的字典。
        2.  **页面设置**: 调用 `apply_page_setup` 应用全局页面设置。
        3.  **内容生成循环**:
            a. 遍历JSON中的每个 `section` 和其中的 `element`。
            b. **预扫描**: 在处理每个节之前，先进行一次预扫描，找出所有带 `bookmark_id` 的段落，并将其ID和文本内容存入一个映射表 `bookmark_map`。
            c. **正式生成**: 再次遍历元素。
               - 对于简单元素（表格、列表、图片等），直接调用相应的 `add_*` 辅助函数。
               - 对于段落，处理其内部的 `content` 数组。当遇到 `text` run时添加文本；当遇到 `footnote`/`endnote` run时，在文档中插入一个唯一的**占位符**字符串，并将占位符和脚注/尾注文本存入 `footnotes_to_inject`/`endnotes_to_inject` 字典；当遇到 `cross_reference` run时，从 `bookmark_map` 中查找目标文本，并调用 `add_cross_reference_field` 插入交叉引用。
               - 对于公式，调用 `get_formula_xml_and_placeholder`，同样在文档中插入占位符，并将占位符和OMML XML存入 `formulas_to_inject` 字典。
        4.  **初始保存**: 将包含所有占位符的 `python-docx` 对象保存到一个内存中的字节流 `processed_bytes`。
        5.  **后处理链 (Post-processing Chain)**:
            a. **公式注入**: 解压 `processed_bytes`，在 `document.xml` 中查找公式占位符，并将其所在的整个 `<w:p>` 元素替换为公式的OMML XML (`<m:oMathPara>`)，然后重新打包。
            b. **脚注/尾注注入**: 将上一步的结果（或初始字节流）传递给 `post_process_footnotes` 和 `post_process_endnotes` 函数，注入脚注和尾注。
            c. **编号格式化**: 将上一步的结果传递给 `apply_numbering_formats`，应用自定义的编号样式。
        6.  **返回结果**: 返回最终处理完毕的 `processed_bytes` 和一份用于调试的XML日志。
    """