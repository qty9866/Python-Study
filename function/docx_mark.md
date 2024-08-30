## 1. Document 类
Document 类是 python-docx 包的核心类，用于创建一个新的 Word 文档或打开一个现有的文档。

**主要方法**

- `Document():` 创建一个新的文档对象，或打开一个现有的文档。
    ```py
    from docx import Document
    # 创建新文档
    doc = Document()
    # 打开现有文档
    doc = Document('existing_document.docx')
    ```

- `save(path)`: 将文档保存到指定路径。
    ```py
    doc.save('new_document.docx')
    ```

- `add_paragraph(text, style=None)`: 添加一个段落到文档中。
    ```py
    paragraph = doc.add_paragraph('This is a new paragraph.')

    ```
- `add_heading(text, level=1)`: 添加标题，level 指定标题级别（1-9）
    ```py
    doc.add_heading('This is a heading', level=1)
    ```
- `add_page_break()`: 添加分页符。
    ```py
    doc.add_page_break()
    ```
- `add_table(rows, cols)`: 添加表格。
    ```py
    table = doc.add_table(rows=3, cols=3)
    ```

## 2. Paragraph 类

`Paragraph` 类表示文档中的一个段落。

**主要方法和属性**

- `text`: 获取或设置段落的文本内容。
    ```py
    paragraph.text = 'Updated paragraph text.'
    ```

- `style`: 获取或设置段落的样式。
    ```py
    paragraph.style = 'Heading 1'
    ```
- `add_run(text)`: 向段落中添加文本块（Run 对象）。
    ```py
    run = paragraph.add_run('This is some more text.')
    ```

## 3. Run 类
`Run` 类表示段落中的一个文本块，可以具有不同的格式。

**主要方法和属性**

- `text`: 获取或设置 Run 的文本内容。
    ```py
    run.text = 'Updated run text.'
    ```
- `bold`: 获取或设置加粗样式。
    ```py
    run.bold = True
    ```
- `italic`: 获取或设置斜体样式。
    ```py
    run.italic = True
    ```
- `underline`: 获取或设置下划线。
    ```py
    run.underline = True
    ```

## 4. Table 类

`Table` 类表示一个表格。

**Table 类表示一个表格。**

- `rows`: 返回表格的行列表。
    ```py
    for row in table.rows:
    pass
    ```
- `columns`: 返回表格的列列表。
    ```py
    for col in table.columns:
    pass
    ```
- `cell(row_idx, col_idx)`: 返回指定单元格对象。
    ```py
    cell = table.cell(0, 0)
    ```

## 5. Cell 类

`Cell` 类表示表格中的一个单元格

**主要方法和属性**

- `text`: 获取或设置单元格中的文本内容。
    ```py
    cell.text = 'Cell content'
    ```
- `merge(other_cell)`: 合并两个单元格。
    ```py
    cell.merge(other_cell)
    ```

## 6. Section 类

`Section` 类表示文档中的一个节。

**主要方法和属性**'

- `start_type`: 设置节的起始类型，如 WD_SECTION.NEW_PAGE。
    ```py
    section.start_type = WD_SECTION.NEW_PAGE
    ```
- `orientation`: 设置节的页面方向，如 WD_ORIENTATION.LANDSCAPE。
    ```py
    section.orientation = WD_ORIENTATION.LANDSCAPE
    ```
- `page_width`: 获取或设置页面宽度。
    ```py
    section.page_width = Inches(8.5)
    ```
- `page_height`: 获取或设置页面高度。
    ```py
    section.page_height = Inches(11)
    ```

## 7. 其他工具和模块

- `shared.Pt` 和 `shared.Inches`: 用于设置字体大小和页面尺寸。
    ```py
    from docx.shared import Pt, Inches
    run.font.size = Pt(12)
    section.page_width = Inches(8.5)
    ```
- `enum.text.WD_PARAGRAPH_ALIGNMENT`: 用于设置段落对齐方式。
    ```py
    from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
    paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    ```
- `enum.text.WD_LINE_SPACING`: 用于设置行距。
    ```py
    from docx.enum.text import WD_LINE_SPACING
    paragraph.paragraph_format.line_spacing = WD_LINE_SPACING.ONE_POINT_FIVE
    ```
- `document.add_picture()`: 用于向文档中添加图片。
    ```py
    doc.add_picture('image.jpg', width=Inches(1.0))
    ```