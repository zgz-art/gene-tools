import streamlit as st
import tempfile
import os
import json
import re
import pdfplumber
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Font, Alignment, Border
from typing import Dict, Any, List, Tuple
from openai import OpenAI

# ----------------------------- 页面配置 -----------------------------
st.set_page_config(page_title="智能简历生成器 (智谱AI)", layout="wide")
st.title("📄 智能简历生成器")
st.markdown("上传 PDF 简历和 Excel 模板，AI 自动填充并**保留原格式**，支持动态增加行。")

# ----------------------------- 1. 模板结构扫描（支持模糊匹配） -----------------------------
def normalize_label(text: str) -> str:
    """将标签文本规范化：去除空格、冒号、换行等，只保留中文字符和字母数字"""
    if not isinstance(text, str):
        return ""
    cleaned = re.sub(r'[：:\s\n\r\t]', '', text)
    return cleaned.strip()

def scan_template_by_keywords(template_path: str) -> Dict[str, Any]:
    wb = openpyxl.load_workbook(template_path, data_only=True)
    ws = wb.worksheets[0]
    
    basic_fields = ['姓名', '身份证号', '出生日期', '电话', '首次参加工作时间', 
                    '首次参加IT领域工作时间', '最高学历', '掌握语言', '掌握技能', '专业证书']
    field_positions = {}
    for row in range(1, ws.max_row + 1):
        for col in range(1, ws.max_column + 1):
            val = ws.cell(row, col).value
            if val and isinstance(val, str):
                norm_val = normalize_label(val)
                for field in basic_fields:
                    if norm_val == field:
                        target_row, target_col = row, col + 1
                        target_cell = ws.cell(row, col + 1)
                        for merged in ws.merged_cells.ranges:
                            if target_cell.coordinate in merged:
                                target_row, target_col = merged.min_row, merged.min_col
                                break
                        field_positions[field] = (target_row, target_col)
                        break
    
    edu_subfields = ['入学时间', '毕业院校', '毕业时间', '专业', '毕业证编号', 
                     '毕业证学信网在线验证码', '学位证编号', '学位证学信网在线验证码']
    edu_positions = {'本科': {}, '硕士': {}}
    
    def find_edu_section(title_keyword: str):
        for row in range(1, ws.max_row + 1):
            val = ws.cell(row, 1).value
            if val and title_keyword in normalize_label(val):
                positions = {}
                for r in range(row + 1, min(row + 20, ws.max_row + 1)):
                    for col in range(1, ws.max_column + 1):
                        v = ws.cell(r, col).value
                        if v and isinstance(v, str):
                            norm_v = normalize_label(v)
                            for sub in edu_subfields:
                                if norm_v == sub:
                                    tr, tc = r, col + 1
                                    target_cell = ws.cell(r, col + 1)
                                    for merged in ws.merged_cells.ranges:
                                        if target_cell.coordinate in merged:
                                            tr, tc = merged.min_row, merged.min_col
                                            break
                                    positions[sub] = (tr, tc)
                                    break
                return positions
        return {}
    
    edu_positions['本科'] = find_edu_section('本科学历')
    edu_positions['硕士'] = find_edu_section('研究生学历')
    
    work_table = None
    for row in range(1, ws.max_row + 1):
        val = ws.cell(row, 1).value
        if val and '工作经历' in str(val) and '由近及远' in str(val):
            header_row = row + 1
            headers = []
            start_col = None
            for col in range(1, ws.max_column + 1):
                h = ws.cell(header_row, col).value
                if h:
                    headers.append(str(h).strip())
                    if start_col is None:
                        start_col = col
            work_table = {
                'header_row': header_row,
                'headers': headers,
                'start_col': start_col,
                'data_start_row': header_row + 1
            }
            break
    
    project_table = None
    for row in range(1, ws.max_row + 1):
        val = ws.cell(row, 1).value
        if val and '项目经历' in str(val) and '与上述工作经历匹配' in str(val):
            header_row = row + 1
            headers = []
            start_col = None
            for col in range(1, ws.max_column + 1):
                h = ws.cell(header_row, col).value
                if h:
                    headers.append(str(h).strip())
                    if start_col is None:
                        start_col = col
            project_table = {
                'header_row': header_row,
                'headers': headers,
                'start_col': start_col,
                'data_start_row': header_row + 1
            }
            break
    
    return {
        'field_positions': field_positions,
        'edu_positions': edu_positions,
        'work_table': work_table,
        'project_table': project_table,
        'sheet_name': ws.title
    }

# ----------------------------- 2. PDF 文本提取 -----------------------------
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

# ----------------------------- 3. 智谱AI调用封装 -----------------------------
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

# ----------------------------- 4. AI 解析（含日期规范化指令） -----------------------------
def parse_resume_with_llm(resume_text: str, template_info: Dict, llm_client: ZhipuAIClient) -> Dict:
    basic_fields = list(template_info['field_positions'].keys())
    edu_subfields = ['入学时间', '毕业院校', '毕业时间', '专业', '毕业证编号', 
                     '毕业证学信网在线验证码', '学位证编号', '学位证学信网在线验证码']
    work_headers = template_info['work_table']['headers'] if template_info['work_table'] else []
    project_headers = template_info['project_table']['headers'] if template_info['project_table'] else []
    
    date_fields = ['出生日期', '首次参加工作时间', '首次参加IT领域工作时间', 
                   '入学时间', '毕业时间', '工作开始日期', '工作结束日期']
    
    prompt = f"""
你是一个专业的简历解析助手。请根据以下简历文本，提取信息，并以严格的JSON格式输出。

需要提取的信息：

1. 基本字段：{basic_fields}

2. 学历信息（分本科和硕士）：
   子字段：{edu_subfields}
   分别填入 "本科" 和 "硕士" 对象中。若无硕士，所有子字段为空字符串。

3. 工作经历（按由近及远排序）：
   每条记录包含列头：{work_headers}

4. 项目经历（按由近及远排序）：
   每条记录包含列头：{project_headers}

**日期格式要求（非常重要）**：
所有日期字段（{date_fields}）必须输出为 "YYYY-MM-DD" 格式。
- 如果原文只给出年月（如“2019年2月”），则输出 "2019-02-01"。
- 如果只给出年份（如“2020”），则输出 "2020-01-01"。
- 如果已经完整到日，保持原样。
- 如果找不到日期，输出空字符串 ""。

输出格式（严格JSON，不要额外解释）：
{{
    "fields": {{"姓名": "...", "身份证号": "...", ...}},
    "education": {{
        "本科": {{"入学时间": "...", "毕业院校": "...", ...}},
        "硕士": {{...}}
    }},
    "work_experiences": [
        {{"工作开始日期": "...", "工作结束日期": "...", "单位名称": "...", "岗位": "..."}},
        ...
    ],
    "project_experiences": [
        {{"工作开始日期": "...", "工作结束日期": "...", "项目名称": "...", "项目描述": "...", "项目角色": "..."}},
        ...
    ]
}}

简历文本：
{resume_text}
"""
    response = llm_client.call(prompt)
    if response.startswith("```json"):
        response = response[7:]
    if response.endswith("```"):
        response = response[:-3]
    data = json.loads(response.strip())
    data.setdefault('fields', {})
    data.setdefault('education', {'本科': {}, '硕士': {}})
    data.setdefault('work_experiences', [])
    data.setdefault('project_experiences', [])
    return data

# ----------------------------- 5. 后处理：日期格式化 -----------------------------
def normalize_date(date_str: str) -> str:
    if not date_str or not isinstance(date_str, str):
        return ""
    date_str = date_str.strip()
    if re.match(r'^\d{4}-\d{2}-\d{2}$', date_str):
        return date_str
    if re.match(r'^\d{4}/\d{1,2}/\d{1,2}$', date_str):
        parts = date_str.split('/')
        y, m, d = parts[0], parts[1].zfill(2), parts[2].zfill(2)
        return f"{y}-{m}-{d}"
    if re.match(r'^\d{4}-\d{1,2}$', date_str):
        y, m = date_str.split('-')
        return f"{y}-{m.zfill(2)}-01（请确认具体日期）"
    if re.match(r'^\d{4}/\d{1,2}$', date_str):
        y, m = date_str.split('/')
        return f"{y}-{m.zfill(2)}-01（请确认具体日期）"
    m1 = re.match(r'(\d{4})年(\d{1,2})月(\d{1,2})日', date_str)
    if m1:
        y, m, d = m1.groups()
        return f"{y}-{m.zfill(2)}-{d.zfill(2)}"
    m2 = re.match(r'(\d{4})年(\d{1,2})月', date_str)
    if m2:
        y, m = m2.groups()
        return f"{y}-{m.zfill(2)}-01（请确认具体日期）"
    if re.match(r'^\d{4}$', date_str):
        return f"{date_str}-01-01（请确认具体日期）"
    return f"{date_str}（请确认日期格式）"

def format_dates_in_data(parsed_data: Dict) -> Dict:
    date_field_names = ['出生日期', '首次参加工作时间', '首次参加IT领域工作时间', 
                        '入学时间', '毕业时间', '工作开始日期', '工作结束日期']
    def process(obj):
        if isinstance(obj, dict):
            for k, v in obj.items():
                if k in date_field_names and isinstance(v, str):
                    obj[k] = normalize_date(v)
                else:
                    process(v)
        elif isinstance(obj, list):
            for item in obj:
                process(item)
    process(parsed_data)
    return parsed_data

# ----------------------------- 6. 填充 Excel（保留格式，动态插入行） -----------------------------
def fill_template_with_format(template_path: str, parsed_data: Dict, output_path: str, template_info: Dict):
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
        positions = template_info['edu_positions'].get(level, {})
        for sub, (row, col) in positions.items():
            if sub in edu_data.get(level, {}):
                ws.cell(row=row, column=col, value=edu_data[level][sub])
    
    # 工作经历（动态行）
    work = template_info['work_table']
    if work and parsed_data.get('work_experiences'):
        start_row = work['data_start_row']
        headers = work['headers']
        start_col = work['start_col']
        records = parsed_data['work_experiences']
        # 计算现有数据行数
        existing = 0
        r = start_row
        while r <= ws.max_row:
            empty = True
            for c in range(start_col, start_col + len(headers)):
                if ws.cell(r, c).value is not None:
                    empty = False
                    break
            if empty:
                break
            existing += 1
            r += 1
        need = len(records) - existing
        if need > 0:
            insert_pos = start_row + existing
            template_row = start_row + existing - 1
            if template_row >= start_row:
                for _ in range(need):
                    ws.insert_rows(insert_pos)
                    for c in range(start_col, start_col + len(headers)):
                        src = ws.cell(template_row, c)
                        dst = ws.cell(insert_pos, c)
                        if src.has_style:
                            dst.font = src.font.copy()
                            dst.fill = src.fill.copy()
                            dst.border = src.border.copy()
                            dst.alignment = src.alignment.copy()
                            dst.number_format = src.number_format
                    insert_pos += 1
        # 写入数据
        for i, rec in enumerate(records):
            cur_row = start_row + i
            for col_idx, header in enumerate(headers):
                if header in rec:
                    ws.cell(row=cur_row, column=start_col + col_idx, value=rec[header])
    
    # 项目经历（同理）
    proj = template_info['project_table']
    if proj and parsed_data.get('project_experiences'):
        start_row = proj['data_start_row']
        headers = proj['headers']
        start_col = proj['start_col']
        records = parsed_data['project_experiences']
        existing = 0
        r = start_row
        while r <= ws.max_row:
            empty = True
            for c in range(start_col, start_col + len(headers)):
                if ws.cell(r, c).value is not None:
                    empty = False
                    break
            if empty:
                break
            existing += 1
            r += 1
        need = len(records) - existing
        if need > 0:
            insert_pos = start_row + existing
            template_row = start_row + existing - 1
            if template_row >= start_row:
                for _ in range(need):
                    ws.insert_rows(insert_pos)
                    for c in range(start_col, start_col + len(headers)):
                        src = ws.cell(template_row, c)
                        dst = ws.cell(insert_pos, c)
                        if src.has_style:
                            dst.font = src.font.copy()
                            dst.fill = src.fill.copy()
                            dst.border = src.border.copy()
                            dst.alignment = src.alignment.copy()
                            dst.number_format = src.number_format
                    insert_pos += 1
        for i, rec in enumerate(records):
            cur_row = start_row + i
            for col_idx, header in enumerate(headers):
                if header in rec:
                    ws.cell(row=cur_row, column=start_col + col_idx, value=rec[header])
    
    wb.save(output_path)

# ----------------------------- 7. Streamlit 界面 -----------------------------
with st.sidebar:
    st.header("🔑 智谱AI 配置")
    api_key = st.text_input("API Key", type="password", help="从 https://bigmodel.cn 获取")
    model_name = st.selectbox("模型", ["glm-4-flash", "glm-4-plus"], index=0)
    st.markdown("---")
    st.markdown("**模板要求**：")
    st.markdown("- 基本字段如 `姓名`、`身份证号` 等（支持冒号和空格）")
    st.markdown("- 学历标题 `本科学历` 和 `研究生学历`")
    st.markdown("- 工作经历标题 `工作经历（由近及远，仅限IT相关经历）`")
    st.markdown("- 项目经历标题 `项目经历（与上述工作经历匹配，仅IT相关经历）`")

col1, col2 = st.columns(2)
with col1:
    pdf_file = st.file_uploader("📄 上传 PDF 简历", type=["pdf"])
with col2:
    template_file = st.file_uploader("📊 上传 Excel 模板", type=["xlsx"])

if st.button("🚀 开始解析并生成", type="primary"):
    if not pdf_file or not template_file:
        st.error("请同时上传 PDF 和 Excel 模板")
        st.stop()
    if not api_key:
        st.error("请填写智谱AI API Key")
        st.stop()
    
    progress = st.progress(0)
    with st.spinner("分析模板结构..."):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp.write(template_file.getbuffer())
            tmp_template_path = tmp.name
        try:
            template_info = scan_template_by_keywords(tmp_template_path)
            st.success("模板识别成功")
            with st.expander("查看识别结果"):
                st.write("基本字段位置：", list(template_info['field_positions'].keys()))
                st.write("本科学历字段：", list(template_info['edu_positions']['本科'].keys()))
                st.write("硕士学历字段：", list(template_info['edu_positions']['硕士'].keys()))
                if template_info['work_table']:
                    st.write("工作经历表头：", template_info['work_table']['headers'])
                if template_info['project_table']:
                    st.write("项目经历表头：", template_info['project_table']['headers'])
        except Exception as e:
            st.error(f"模板扫描失败：{e}")
            st.stop()
    progress.progress(20)
    
    with st.spinner("读取 PDF..."):
        try:
            resume_text = extract_text_from_pdf(pdf_file)
            if not resume_text:
                st.error("PDF 内容为空")
                st.stop()
        except Exception as e:
            st.error(f"PDF读取失败：{e}")
            st.stop()
    progress.progress(40)
    
    with st.spinner(f"调用智谱AI ({model_name}) 解析..."):
        try:
            client = ZhipuAIClient(api_key, model_name)
            parsed_raw = parse_resume_with_llm(resume_text, template_info, client)
            parsed_data = format_dates_in_data(parsed_raw)
            st.success("解析完成")
            with st.expander("查看提取的数据（日期已标准化）"):
                st.json(parsed_data)
        except Exception as e:
            st.error(f"AI解析失败：{e}")
            st.stop()
    progress.progress(70)
    
    with st.spinner("生成Excel文件..."):
        try:
            output_temp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx").name
            fill_template_with_format(tmp_template_path, parsed_data, output_temp, template_info)
            with open(output_temp, "rb") as f:
                excel_data = f.read()
            st.download_button("📥 下载生成的简历", data=excel_data, file_name="filled_resume.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            os.unlink(tmp_template_path)
            os.unlink(output_temp)
        except Exception as e:
            st.error(f"生成失败：{e}")
            st.stop()
    progress.progress(100)
    st.balloons()
