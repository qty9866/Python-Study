import win32com.client as win32
from docx import Document
import re
import os


def is_chinese_date_format(date_text):
    pattern = r'[〇一二三四五六七八九十]{4}年[一二三四五六七八九十]{1,2}月'
    match = re.match(pattern, date_text)
    return match is not None


def mark_incorrect_date(paragraph, incorrect_text):
    start = paragraph.Range.Text.find(incorrect_text)
    # end = start + len(incorrect_text)
    rng = paragraph.Range.Duplicate
    rng.Start += start
    rng.End = rng.Start + len(incorrect_text)
    rng.Font.Color = win32.constants.wdColorRed
    comment = paragraph.Range.Comments.Add(rng, "日期格式错误")
    return comment


# 首页日期格式检查
def date_format_check(doc):
    first_page = doc.Paragraphs
    for paragraph in first_page:
        text = paragraph.Range.Text.strip()
        match = re.search(r'\d{4}年\d{1,2}月|\d{4}年[一二三四五六七八九十]{1,2}月|[〇一二三四五六七八九十]{4}年\d{1,2}月', text)
        if match:
            date = match.group(0)
            if not is_chinese_date_format(date):
                mark_incorrect_date(paragraph, date)
            break


# def extract_person_info(text):
#     zong_name_match = re.search(r'项目总负责人：([^\n]+)', text)
#     dan_name_match = re.search(r'单项设计负责人：([^\n]+)', text)
#     jian_name_match = re.search(r'建设单位联系人：([^\n]+)', text)
#     tel_match = re.findall(r'电话：([^\n]+)', text)
#     mail_match = re.findall(r'电子邮箱：([^\n]+)', text)
#     persons = []
#     if zong_name_match and len(tel_match) > 0 and len(mail_match) > 0:
#         person1 = {
#             'name': zong_name_match.group(1).strip(),
#             'tel': tel_match[0].strip(),
#             'mail': mail_match[0].strip()
#         }
#         persons.append(person1)
#     if dan_name_match and len(tel_match) > 1 and len(mail_match) > 1:
#         person2 = {
#             'name': dan_name_match.group(1).strip(),
#             'tel': tel_match[1].strip(),
#             'mail': mail_match[1].strip()
#         }
#         persons.append(person2)
#     if jian_name_match and len(tel_match) > 2 and len(mail_match) > 2:
#         person3 = {
#             'name': jian_name_match.group(1).strip(),
#             'tel': tel_match[2].strip(),
#             'mail': mail_match[2].strip()
#         }
#         persons.append(person3)
#     return persons if persons else None


def extract_person_info(docx_path):
    docx = Document(docx_path)
    paragraph_text = '设计文件分发表'
    found_paragraph = False
    for para in docx.paragraphs:
        if paragraph_text in para.text:
            for table in docx.tables:
                if para._element.getnext() is table._element:
                    last_row = table.rows[-1]
                    last_cell = last_row.cells[-1]
                    text = last_cell.text.strip()
                    zong_name_match = re.search(r'项目总负责人：([^\n]+)', text)
                    dan_name_match = re.search(r'单项设计负责人：([^\n]+)', text)
                    jian_name_match = re.search(r'建设单位联系人：([^\n]+)', text)
                    tel_match = re.findall(r'电话：([^\n]+)', text)
                    mail_match = re.findall(r'电子邮箱：([^\n]+)', text)
                    persons = []
                    if zong_name_match and len(tel_match) > 0 and len(mail_match) > 0:
                        person1 = {
                            'name': zong_name_match.group(1).strip(),
                            'tel': tel_match[0].strip(),
                            'mail': mail_match[0].strip()
                        }
                        persons.append(person1)
                    if dan_name_match and len(tel_match) > 1 and len(mail_match) > 1:
                        person2 = {
                            'name': dan_name_match.group(1).strip(),
                            'tel': tel_match[1].strip(),
                            'mail': mail_match[1].strip()
                        }
                        persons.append(person2)
                    if jian_name_match and len(tel_match) > 2 and len(mail_match) > 2:
                        person3 = {
                            'name': jian_name_match.group(1).strip(),
                            'tel': tel_match[2].strip(),
                            'mail': mail_match[2].strip()
                        }
                        persons.append(person3)
                    return persons if persons else None
    print("未找到设计文件分发表")
    return None

def check_paragraph_format(paragraph):
    """
    检查段落中的文本格式是否符合要求：
    - 中文字体：宋体
    - 英文字体：Times New Roman
    - 字体大小：小四（12pt）
    - 行距：1.5倍
    如果不满足要求，则返回具体不满足的条件，否则返回None。
    """
    format_issues = set()

    # 检查段落的行距是否为1.5倍行距
    if paragraph.LineSpacingRule != win32.constants.wdLineSpace1pt5:
        format_issues.add("行距不为1.5倍")

    # 遍历段落中的每个字符，检查字体和字号
    for char in paragraph.Range.Characters:
        font_name = char.Font.Name
        font_size = char.Font.Size
        text = char.Text

        # 判断中文字符的字体
        if '\u4e00' <= text <= '\u9fff':  # 中文字符范围
            if font_name != '宋体':
                format_issues.add("中文字体不是宋体")
            if font_size != 12:  # 小四对应的大小
                format_issues.add("中文字体大小不是小四")

        # 判断英文字符和数字的字体
        elif text.isalpha() or text.isdigit():  # 英文或数字
            if font_name != 'Times New Roman':
                format_issues.add("英文或数字字体不是Times New Roman")
            if font_size != 12:
                format_issues.add("英文或数字字体大小不是小四")

    return list(format_issues) if format_issues else None

def add_comment(paragraph, issues):
    rng = paragraph.Range.Duplicate
    comment_text = "未满足的格式要求：" + "；".join(issues)
    paragraph.Range.Comments.Add(rng, comment_text)


if __name__ == '__main__':
    # parser = argparse.ArgumentParser()
    # parser.add_argument('f', help='word file path')
    # args = parser.parse_args()

    word_app = win32.gencache.EnsureDispatch('Word.Application')
    word_app.Visible = False  # 不显示 Word 窗口

    doc_path = 'test.docx'

    # doc = word_app.Documents.Open(os.path.abspath(doc_path))


    # date_format_check(doc)
    persons = extract_person_info(doc_path)     # 提取分发表人员信息
    print(persons)


    # doc.SaveAs(os.path.abspath('test_processed.docx'))
    # doc.Close()
    # word_app.Quit()
