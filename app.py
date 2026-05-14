import streamlit as st
import tempfile
import os
import json
import pdfplumber
import openpyxl
from openpyxl.utils import get_column_letter
from typing import Dict, Any, List, Tuple
import re
from openai import OpenAI
from datetime import datetime

# ----------------------------- 页面配置 -----------------------------
st.set_page_config(page_title="简历解析器 (自定义模板)", layout="wide")
st.title("📄 简历解析 → 指定模板填充")
st.markdown("根据预设的关键字段和表格区域，自动提取并填充。所有日期格式为 YYYY-MM-DD，缺失日则默认为01日并添加提醒。")

# ----------------------------- 日期格式化函数 -----------------------------
def format_date_string(date_str: str) -> str:
    """将各种日期格式统一转为 YYYY-MM-DD，若无法解析则返回原字符串+提醒"""
    if not date_str or not isinstance(date_str, str):
        return date_str
    # 尝试匹配常见格式
    # 格式1: YYYY-MM-DD 或 YYYY/MM/DD
    match = re.match(r'(\d{4})[-/](\d{1,2})[-/](\d{1,2})', date_str)
    if match:
        y, m, d = match.groups()
        y, m, d = int(y), int(m), int(d)
        if m < 1: m = 1
        if m > 12: m = 12
        if d < 1: d = 1
        max_days = [31, 29 if (y % 4 == 0 and y % 100 != 0) or (y % 400 == 0) else 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][m-1]
        if d > max_days:
            d = max_days
        formatted = f"{y:04d}-{m:02d}-{d:02d}"
        return formatted
    # 格式2: YYYY-MM 或 YYYY/MM
    match = re.match(r'(\d{4})[-/](\d{1,2})$', date_str)
    if match:
        y, m = match.groups()
        y, m = int(y), int(m)
        if m < 1: m = 1
        if m > 12: m = 12
        formatted = f"{y:04d}-{m:02d}-01"
        return formatted + "（请确认具体日期）"
    # 格式3: 只写年份
    match = re.match(r'^(\d{4})$', date_str.strip())
    if match:
        y = int(match.group(1))
        formatted = f"{y:04d}-01-01"
        return formatted + "（请确认具体日期）"
    # 其他情况
    return f"{date_str}（请确认日期格式）"

def format_dates_in_data(data: Dict) -> Dict:
    """递归处理所有字符串值，对日期字段进行格式化"""
    date_fields = [
        '出生日期', '首次参加工作时间', '首次参加IT领域工作时间',
        '入学时间', '毕业时间', '工作开始日期', '工作结束日期'
    ]
    def process_value(key, value):
        if isinstance(value, str) and any(df in key for df in date_fields):
            return format_date_string(value)
        return value
    def process_dict(d):
        if not isinstance(d, dict):
            return d
        new_d = {}
        for k, v in d.items():
            if isinstance(v, dict):
                new_d[k] = process_dict(v)
            elif isinstance(v, list):
                new_d[k] = [process_dict(item) if isinstance(item, dict) else process_value(k, item) for item in v]
            else:
                new_d[k] = process_value(k, v)
        return new_d
    return process_dict(data)

# ----------------------------- 模板结构扫描（基于关键词） -----------------------------
def scan_template_by_keywords(template_path: str) -> Dict[str, Any]:
    wb = openpyxl.load_workbook(template_path, data_only=True)
    ws = wb.worksheets[0]
    
    field_positions = {}
    basic_fields = ['姓名', '身份证号', '出生日期', '电话', '首次参加工作时间', '首次参加IT领域工作时间', '最高学历', '掌握语言', '掌握技能', '专业证书']
    
    for row in range(1, ws.max_row + 1):
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row, col)
            if cell.value and isinstance(cell.value, str):
                val = cell.value.strip()
                for field in basic_fields:
                    if val.startswith(field) or val.startswith(field + '：') or val.startswith(field + ':'):
                        target_cell = ws.cell(row, col + 1)
                        target_row, target_col = row, col + 1
                        for merged in ws.merged_cells.ranges:
                            if target_cell.coordinate in merged:
                                target_row, target_col = merged.min_row, merged.min_col
                                break
                        field_positions[field] = (target_row, target_col)
                        break
    
    edu_subfields = ['入学时间', '毕业院校', '毕业时间', '专业', '毕业证编号', '毕业证学信网在线验证码', '学位证编号', '学位证学信网在线验证码']
    education_positions = {'本科': {}, '硕士': {}}
    
    def find_edu_section(title_keyword: str):
        for row in range(1, ws.max_row + 1):
            cell = ws.cell(row, 1)
            if cell.value and title_keyword in str(cell.value):
                positions = {}
                for r in range(row + 1, min(row + 30, ws.max_row + 1)):
                    for col in range(1, ws.max_column + 1):
                        val = ws.cell(r, col).value
                        if val and isinstance(val, str):
                            for sub in edu_subfields:
                                if val.startswith(sub) or val.startswith(sub + '：') or val.startswith(sub + ':'):
                                    target_cell = ws.cell(r, col + 1)
                                    tr, tc = r, col + 1
                                    for merged in ws.merged_cells.ranges:
                                        if target_cell.coordinate in merged:
                                            tr, tc = merged.min_row, merged.min_col
                                            break
                                    positions[sub] = (tr, tc)
                                    break
                return positions
        return {}
    
    education_positions['本科'] = find_edu_section('本科学历')
    education_positions['硕士'] = find_edu_section('研究生学历')
    
    work_table = None
    project_table = None
    
    for row in range(1, ws.max_row + 1):
        cell = ws.cell(row, 1)
        if cell.value and '工作经历' in str(cell.value) and '由近及远' in str(cell.value):
            header_row = row + 1
            headers = []
            start_col = None
            for col in range(1, ws.max_column + 1):
                val = ws.cell(header_row, col).value
                if val:
                    headers.append(str(val).strip())
                    if start_col is None:
                        start_col = col
            work_table = {
                'header_row': header_row,
                'headers': headers,
                'data_start_row': header_row + 1,
                'start_col': start_col
            }
        elif cell.value and '项目经历' in str(cell.value) and '与上述工作经历匹配' in str(cell.value):
            header_row = row + 1
            headers = []
            start_col = None
            for col in range(1, ws.max_column + 1):
                val = ws.cell(header_row, col).value
                if val:
                    headers.append(str(val).strip())
                    if start_col is None:
                        start_col = col
            project_table = {
                'header_row': header_row,
                'headers': headers,
                'data_start_row': header_row + 1,
                'start_col': start_col
            }
    
    return {
        'field_positions': field_positions,
        'education_positions': education_positions,
        'work_table': work_table,
        'project_table': project_table,
        'sheet_name': ws.title
    }

# ----------------------------- PDF文本提取 -----------------------------
def extract_text_from_pdf(pdf_file) -> str:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(pdf_file.getbuffer())
        tmp_path = tmp.name
    text = ""
    with pdfplumber.open(tmp_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n"
    os.unlink(tmp_path)
    return text.strip()

# ----------------------------- 智谱AI客户端 -----------------------------
class ZhipuAIClient:
    def __init__(self, api_key: str, model_name: str = "glm-4-flash"):
        self.client = OpenAI(api_key=api_key, base_url="https://open.bigmodel.cn/api/paas/v4/")
        self.model_name = model_name
    def call(self, prompt: str) -> str:
        response = self.client.chat.completions.create(
            model=self.model_name,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.0
        )
        return response.choices[0].message.content

# ----------------------------- AI解析（含日期格式要求） -----------------------------
def parse_resume_with_llm(resume_text: str, template_info: Dict, llm_client: ZhipuAIClient) -> Dict:
    basic_fields = list(template_info['field_positions'].keys())
    edu_subfields = ['入学时间', '毕业院校', '毕业时间', '专业', '毕业证编号', '毕业证学信网在线验证码', '学位证编号', '学位证学信网在线验证码']
    work_headers = template_info['work_table']['headers'] if template_info['work_table'] else []
    project_headers = template_info['project_table']['headers'] if template_info['project_table'] else []
    
    prompt = f"""
你是一个专业的简历解析助手。请根据以下简历文本，提取信息，并以严格的JSON格式输出。

需要提取的信息：

1. 基本字段：{basic_fields}

2. 学历信息（分本科和硕士，如果没有硕士则留空）：
   子字段：{edu_subfields}
   请分别填入 "本科" 和 "硕士" 对象中。

3. 工作经历（按由近及远排序）：
   每条记录包含列头：{work_headers}

4. 项目经历（按由近及远排序）：
   每条记录包含列头：{project_headers}

**重要日期格式要求**：
- 所有涉及时间的字段（出生日期、首次参加工作时间、首次参加IT领域工作时间、入学时间、毕业时间、工作开始日期、工作结束日期）都必须以 "YYYY-MM-DD" 格式输出。
- 如果简历中只提供了年份和月份（如“2019年2月”），则输出 "2019-02-01"。
- 如果只提供了年份（如“2019年”），则输出 "2019-01-01"。
- 如果完全没有提供日期信息，输出空字符串 ""。
- 不要添加任何额外说明文字，只输出日期字符串。

输出格式（严格JSON，不要额外解释）：
{{
    "fields": {{"姓名": "值", "身份证号": "值", "出生日期": "1990-01-01", ...}},
    "education": {{
        "本科": {{"入学时间": "2016-09-01", "毕业时间": "2020-06-01", ...}},
        "硕士": {{...}}
    }},
    "work_experiences": [
        {{"工作开始日期": "2020-01-01", "工作结束日期": "2023-12-01", "单位名称": "...", "岗位": "..."}},
        ...
    ],
    "project_experiences": [
        {{"工作开始日期": "2021-03-01", "工作结束日期": "2021-08-01", "项目名称": "...", "项目描述": "...", "项目角色": "..."}},
        ...
    ]
}}

注意：
- 如果找不到信息，填写空字符串。
- 工作经历和项目经历按时间由近及远排列（最近的在前）。
- 所有日期必须严格遵循 YYYY-MM-DD 格式，缺失日期的部分用01补全。

简历文本：
{resume_text}
"""
    response = llm_client.call(prompt)
    if response.startswith("```json"):
        response = response[7:]
    if response.endswith("```"):
        response = response[:-3]
    data = json.loads(response.strip())
    if 'fields' not in data:
        data['fields'] = {}
    if 'education' not in data:
        data['education'] = {'本科': {}, '硕士': {}}
    if 'work_experiences' not in data:
        data['work_experiences'] = []
    if 'project_experiences' not in data:
        data['project_experiences'] = []
    
    # 后处理日期格式
    data = format_dates_in_data(data)
    return data

# ----------------------------- 填充Excel（保持格式） -----------------------------
def fill_template(template_path: str, parsed_data: Dict, output_path: str, template_info: Dict):
    wb = openpyxl.load_workbook(template_path)
    ws = wb[template_info['sheet_name']]
    
    # 基本字段
    fields = parsed_data.get('fields', {})
    for field, (row, col) in template_info['field_positions'].items():
        if field in fields:
            ws.cell(row=row, column=col, value=fields[field])
    
    # 学历
    edu_data = parsed_data.get('education', {})
    for level in ['本科', '硕士']:
        if level in edu_data:
            positions = template_info['education_positions'].get(level, {})
            for subfield, (row, col) in positions.items():
                if subfield in edu_data[level]:
                    ws.cell(row=row, column=col, value=edu_data[level][subfield])
    
    # 工作经历
    work_table = template_info['work_table']
    if work_table and parsed_data.get('work_experiences'):
        start_row = work_table['data_start_row']
        # 删除原有数据行
        end_row = start_row
        while end_row <= ws.max_row:
            row_is_empty = True
            for col in range(work_table['start_col'], work_table['start_col'] + len(work_table['headers'])):
                if ws.cell(end_row, col).value is not None:
                    row_is_empty = False
                    break
            if row_is_empty:
                break
            end_row += 1
        if end_row > start_row:
            ws.delete_rows(start_row, end_row - start_row)
        for i, record in enumerate(parsed_data['work_experiences']):
            current_row = start_row + i
            for col_idx, header in enumerate(work_table['headers']):
                if header in record:
                    ws.cell(row=current_row, column=work_table['start_col'] + col_idx, value=record[header])
    
    # 项目经历
    project_table = template_info['project_table']
    if project_table and parsed_data.get('project_experiences'):
        start_row = project_table['data_start_row']
        end_row = start_row
        while end_row <= ws.max_row:
            row_is_empty = True
            for col in range(project_table['start_col'], project_table['start_col'] + len(project_table['headers'])):
                if ws.cell(end_row, col).value is not None:
                    row_is_empty = False
                    break
            if row_is_empty:
                break
            end_row += 1
        if end_row > start_row:
            ws.delete_rows(start_row, end_row - start_row)
        for i, record in enumerate(parsed_data['project_experiences']):
            current_row = start_row + i
            for col_idx, header in enumerate(project_table['headers']):
                if header in record:
                    ws.cell(row=current_row, column=project_table['start_col'] + col_idx, value=record[header])
    
    wb.save(output_path)

# ----------------------------- Streamlit界面 -----------------------------
with st.sidebar:
    st.header("🔑 智谱AI 配置")
    api_key = st.text_input("API Key", type="password", help="从 https://bigmodel.cn 获取")
    model_name = st.selectbox("模型", ["glm-4-flash", "glm-4-plus"], index=0)
    st.markdown("---")
    st.markdown("**模板要求**：")
    st.markdown("- 基本字段标签如 `姓名：`、`身份证号：` 等")
    st.markdown("- 学历部分需有 `本科学历` 和 `研究生学历` 标题")
    st.markdown("- 工作经历标题行：`工作经历（由近及远，仅限IT相关经历）`")
    st.markdown("- 项目经历标题行：`项目经历（与上述工作经历匹配，仅IT相关经历）`")

col1, col2 = st.columns(2)
with col1:
    pdf_file = st.file_uploader("📄 上传 PDF 简历", type=["pdf"])
with col2:
    template_file = st.file_uploader("📊 上传 Excel 模板", type=["xlsx"])

if st.button("🚀 开始解析并生成", type="primary"):
    if not pdf_file or not template_file:
        st.error("请同时上传文件")
        st.stop()
    if not api_key:
        st.error("请填写API Key")
        st.stop()
    
    progress = st.progress(0)
    with st.spinner("扫描模板结构..."):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp.write(template_file.getbuffer())
            tmp_template_path = tmp.name
        try:
            template_info = scan_template_by_keywords(tmp_template_path)
            st.success("模板结构识别成功")
            with st.expander("查看识别结果"):
                st.write("基本字段位置：", list(template_info['field_positions'].keys()))
                st.write("本科字段位置：", list(template_info['education_positions']['本科'].keys()))
                st.write("硕士字段位置：", list(template_info['education_positions']['硕士'].keys()))
                if template_info['work_table']:
                    st.write("工作经历表头：", template_info['work_table']['headers'])
                if template_info['project_table']:
                    st.write("项目经历表头：", template_info['project_table']['headers'])
        except Exception as e:
            st.error(f"模板扫描失败：{e}")
            st.stop()
    progress.progress(20)
    
    with st.spinner("读取PDF..."):
        try:
            resume_text = extract_text_from_pdf(pdf_file)
            if not resume_text:
                st.error("PDF内容为空")
                st.stop()
        except Exception as e:
            st.error(f"PDF读取失败：{e}")
            st.stop()
    progress.progress(40)
    
    with st.spinner(f"调用智谱AI ({model_name}) 解析简历..."):
        try:
            client = ZhipuAIClient(api_key, model_name)
            parsed = parse_resume_with_llm(resume_text, template_info, client)
            st.success("解析完成")
            with st.expander("查看提取的数据"):
                st.json(parsed)
        except Exception as e:
            st.error(f"AI解析失败：{e}")
            st.stop()
    progress.progress(70)
    
    with st.spinner("生成Excel文件..."):
        try:
            output_temp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx").name
            fill_template(tmp_template_path, parsed, output_temp, template_info)
            with open(output_temp, "rb") as f:
                excel_data = f.read()
            st.download_button("📥 下载生成的Excel", data=excel_data, file_name="filled_resume.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            os.unlink(tmp_template_path)
            os.unlink(output_temp)
        except Exception as e:
            st.error(f"生成失败：{e}")
            st.stop()
    progress.progress(100)
    st.balloons()
