import xml.etree.ElementTree as ET

namespaces = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
}

def load_xml(file_path):
    """加载并解析XML文件"""
    try:
        tree = ET.parse(file_path)
        root = tree.getroot()
        print("XML文件加载成功")  # 调试信息
        return root
    except Exception as e:
        print(f"XML文件加载失败: {e}")
        return None

def extract_toc_entries(root):
    """提取目录项信息，包括字体和行距"""
    toc_entries = []
    paragraphs = root.findall('.//w:p', namespaces)
    print(f"找到 {len(paragraphs)} 个段落")  # 调试信息

    for paragraph in paragraphs:
        pPr = paragraph.find('.//w:pPr', namespaces)
        if pPr is not None:
            pStyle = pPr.find('.//w:pStyle', namespaces)
            if pStyle is not None and pStyle.attrib.get('{%s}val' % namespaces['w']).startswith('TOC'):
                print("找到目录段落")  # 调试信息
                # 提取字体、字号、行距等信息
                rFonts = pPr.find('.//w:rFonts', namespaces)
                ascii_font = rFonts.attrib.get('{%s}ascii' % namespaces['w']) if rFonts is not None else None
                eastAsia_font = rFonts.attrib.get('{%s}eastAsia' % namespaces['w']) if rFonts is not None else None
                cs_font = rFonts.attrib.get('{%s}cs' % namespaces['w']) if rFonts is not None else None

                sz = pPr.find('.//w:sz', namespaces)
                font_size = sz.attrib.get('{%s}val' % namespaces['w']) if sz is not None else None

                spacing = pPr.find('.//w:spacing', namespaces)
                line_spacing = spacing.attrib.get('{%s}line' % namespaces['w']) if spacing is not None else None

                texts = [node.text for node in paragraph.findall('.//w:t', namespaces) if node.text]
                entry_text = ''.join(texts)

                toc_entries.append({
                    'text': entry_text,
                    'ascii_font': ascii_font,
                    'eastAsia_font': eastAsia_font,
                    'cs_font': cs_font,
                    'font_size': font_size,
                    'line_spacing': line_spacing
                })

    return toc_entries

def check_toc_format(toc_entries):
    """检查目录格式并输出是否符合要求"""
    if not toc_entries:
        print("未找到任何目录项")  # 调试信息
        return

    for entry in toc_entries:
        print(f"目录项: {entry['text']}")
        print(f"  西文字体: {entry['ascii_font']}")
        print(f"  中文字体: {entry['eastAsia_font']}")
        print(f"  CS字体: {entry['cs_font']}")
        print(f"  字号: {entry['font_size']}")
        print(f"  行距: {entry['line_spacing']}")

        if entry['eastAsia_font'] != '宋体':
            print("  错误: 中文字体不符合要求")
        if entry['ascii_font'] != 'Times New Roman':
            print("  错误: 西文字体不符合要求")
        if entry['font_size'] != '24':
            print("  错误: 字号不符合要求")
        if entry['line_spacing'] != '360':
            print("  错误: 行距不符合要求")
        print()

def main(file_path):
    root = load_xml(file_path)
    if root is None:
        return
    toc_entries = extract_toc_entries(root)
    check_toc_format(toc_entries)

if __name__ == "__main__":
    file_path = "function/tt_xml/word/document.xml"
    main(file_path)
