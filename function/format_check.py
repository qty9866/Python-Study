import re
import csv
from docx import Document
from docx.shared import Pt

# 首页日期格式检测
def check_date_format(file_path):
    doc = Document(file_path)
    pattern = r'[〇一二三四五六七八九零]{4}年[〇一二三四五六七八九十]{1,2}月'
    first_page = ''
    for par in doc.paragraphs:
        first_page += par.text + '\n'
        if 'PAGE_BREAK' in par.text:
            break
    match = re.search(pattern, first_page)
    if not match:
        print("首页日期格式有误")

def is_chinese(text):
    for ch in text:
        if '\u4e00' <= ch <= '\u9fff':
            return True
    return False

# 正文格式检测
def check_paragraph_formatting(file_path):
    chinese_font = "宋体"
    english_font = "Times New Roman"
    font_size = Pt(12)  
    line_spacing = 1.5
    doc = Document(file_path)
    
    # 确认段落的样式，只针对Normal样式进行检测
    for paragraph in doc.paragraphs:
        if paragraph.style.name != 'Normal':
            continue
        
        issues = []
        runs = paragraph.runs
        for run in runs:
            text = run.text.strip()
            if not text:
                continue  # 跳过空文本
            
            if is_chinese(text):
                if run.font.name != chinese_font or run.font.size != font_size:
                    issues.append(f"中文字体应为 {chinese_font}，实际为 {run.font.name if run.font.name else '未设置'}，字体大小应为 {font_size.pt}pt，实际为 {run.font.size.pt if run.font.size else '未设置'}pt")
            else:
                if run.font.name != english_font or run.font.size != font_size:
                    issues.append(f"英文/数字字体应为 {english_font}，实际为 {run.font.name if run.font.name else '未设置'}，字体大小应为 {font_size.pt}pt，实际为 {run.font.size.pt if run.font.size else '未设置'}pt")
        
        if paragraph.paragraph_format.line_spacing != line_spacing:
            issues.append(f"行距应为 {line_spacing} 倍，实际为 {paragraph.paragraph_format.line_spacing if paragraph.paragraph_format.line_spacing else '未设置'}")

        if issues:
            print(f"段落内容: {paragraph.text}\n问题: {', '.join(issues)}\n")


def extract_person_info(text):
    zong_name_match = re.search(r'项目总负责人：([^\n]+)', text)
    dan_name_match = re.search(r'单项设计负责人：([^\n]+)',text)
    jian_name_match = re.search(r'建设单位联系人：([^\n]+)',text)
    tel_match = re.findall(r'电话：([^\n]+)', text)
    mail_match = re.findall(r'电子邮箱：([^\n]+)', text)

    persons = []
    if zong_name_match and len(tel_match)>0 and len(mail_match)>0:
        person1={
            'name':zong_name_match.group(1).strip(),
            'tel':tel_match[0].strip(),
            'mail':mail_match[0].strip()
        }
        persons.append(person1)
    if dan_name_match and len(tel_match) > 1 and len(mail_match) > 1:
        person2={
            'name':dan_name_match.group(1).strip(),
            'tel':tel_match[1].strip(),
            'mail':mail_match[1].strip()
        }   
        persons.append(person2)
    if jian_name_match and len(tel_match) > 2 and len(mail_match) > 2 :
        person3={
            'name':jian_name_match.group(1).strip(),
            'tel':tel_match[2].strip(),
            'mail':mail_match[2].strip()
        }
        persons.append(person3)
    return persons if persons else None

def find_table_after_paragraph(docx_path, paragraph_text):
    doc = Document(docx_path)
    # 标记找到指定段落后的第一个表格
    found_paragraph = False
    
    # 遍历所有段落和表格
    for para in doc.paragraphs:
        if paragraph_text in para.text:
            found_paragraph = True
        
        # 如果已经找到了指定段落，检查下一个表格
        if found_paragraph:
            for table in doc.tables:
                if para._element.getnext() is table._element:
                    # 获取表格的最右下角单元格内容
                    last_row = table.rows[-1]
                    last_cell = last_row.cells[-1]
                    return last_cell.text.strip()
    
    # 如果没有找到对应表格，返回空字符串或其他提示信息
    return None

# 使用示例
docx_path = 'test.docx'  # 替换为你的文件路径
paragraph_text = '设计文件分发表'  # 替换为段落中的目标文字

content = find_table_after_paragraph(docx_path, paragraph_text)
# if content:
#     print(content)
# else:
#     print(f"未找到位于'{paragraph_text}'段落后面的表格。")
persons = extract_person_info(content)
print(persons)


def find_person_in_csv(csv_path, persons):
    # 读取 CSV 文件
    with open(csv_path, newline='', encoding='GBK') as csvfile:
        reader = csv.DictReader(csvfile)
        
        # 将 CSV 文件中的数据转换为列表
        data = list(reader)
        
    # 查找并输出不完全符合的人员姓名
    unmatched_persons = []
    for person in persons:
        matched = False
        for row in data:
            if (row['姓名'].strip() == person['name'] and
                row['手机号'].strip() == person['tel'] and
                row['邮箱'].strip() == person['mail']):
                matched = True
                break
        if not matched:
            unmatched_persons.append(person['name'])
    
    return unmatched_persons

# 使用示例
csv_path = 'information.csv'  # 替换为你的CSV文件路径

unmatched = find_person_in_csv(csv_path, persons)
if unmatched:
    print("以下人员的信息在 CSV 文件中不完全符合:")
    for name in unmatched:
        print(name)
else:
    print("所有人员的信息都完全符合。")







# 指定文档路径
doc_path = 'tt.docx'



# 使用示例
check_date_format(doc_path)
check_paragraph_formatting(doc_path)


