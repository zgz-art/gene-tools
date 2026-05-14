import streamlit as st
import tempfile
import os
import json
import pdfplumber
import openpyxl
from openpyxl.utils import get_column_letter, range_boundaries
from typing import Dict, Any, List, Union, Tuple
import re
from openai import OpenAI

# ----------------------------- 页面配置 -----------------------------
st.set_page_config(page_title="智能简历解析 (合并单元格模板)", layout="wide")
st.title("📄 AI 简历解析 → 复杂 Excel 模板")
st.markdown("支持合并单元格、键值对字段、多行表格（如工作经历、项目经历），**保留原格式**。")

# ----------------------------- 辅助函数 -----------------------------
def get_merged_cell_value(ws, cell):
    """获取单元格实际值（处理合并单元格）"""
    for merged_range in ws.merged_cells.ranges:
        if cell.coordinate in merged_range:
            return ws.cell(merged_range.min_row, merged_range.min_col).value
    return cell.value

def scan_template_structure(template_path: str) -> Dict[str, Any]:
    """
    扫描第一个sheet，识别：
    - key_value_fields: {'姓名': (row, col), ...}  用于单个字段填充
    - tables: [{'name': '工作经历', 'header_row': row, 'start_col': col, 'headers': [...], 'data_start_row': row+1}, ...]
    """
    wb = openpyxl.load_workbook(template_path, data_only=True)
    sheet_name = wb.sheetnames[0]
    ws = wb[sheet_name]
    
    # 1. 识别键值对字段（内容以"："或":"结尾，且右侧有可填充单元格）
    key_value_fields = {}
    for row in range(1, ws.max_row + 1):
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row, col)
            value = cell.value
            if value and isinstance(value, str):
                value = value.strip()
                # 匹配如 "姓名："、"姓名:" 或 "姓名：" 带冒号的模式
                if re.search(r'[：:]$', value):
                    label = value.rstrip('：:').strip()
                    # 右侧单元格（或合并单元格）作为填充目标
                    target_cell = ws.cell(row, col + 1)
                    # 如果右侧是合并单元格，找到合并区域的实际位置
                    target_row, target_col = row, col + 1
                    for merged in ws.merged_cells.ranges:
                        if target_cell.coordinate in merged:
                            target_row, target_col = merged.min_row, merged.min_col
                            break
                    key_value_fields[label] = (target_row, target_col)
    
    # 2. 识别表格区域（通过常见表头关键词，如“工作开始日期”、“单位名称”等）
    tables = []
    # 定义表头关键词组
    table_keywords = {
        "工作经历": ["工作开始日期", "工作结束日期", "单位名称", "岗位/职务", "是否邮储银行自主研发工作经验"],
        "项目经历": ["工作开始日期", "工作结束日期", "项目名称", "项目描述", "项目角色", "是否邮储银行自主研发工作经验"]
    }
    for table_name, keywords in table_keywords.items():
        for row in range(1, ws.max_row + 1):
            header_row_cells = [ws.cell(row, col).value for col in range(1, ws.max_column + 1) if ws.cell(row, col).value]
            # 检查该行是否包含所有关键词（或大部分）
            matched = all(any(kw in str(cell) for cell in header_row_cells) for kw in keywords[:2])  # 至少匹配前两个
            if matched:
                # 找到表头各列的位置
                headers = []
                start_col = None
                for col in range(1, ws.max_column + 1):
                    val = ws.cell(row, col).value
                    if val:
                        headers.append(str(val).strip())
                        if start_col is None:
                            start_col = col
                tables.append({
                    'name': table_name,
                    'header_row': row,
                    'start_col': start_col,
                    'headers': headers,
                    'data_start_row': row + 1
                })
                break  # 每个表只找一次
    
    return {
        'key_value_fields': key_value_fields,
        'tables': tables,
        'sheet_name': sheet_name
    }

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

# ----------------------------- 智谱AI调用 -----------------------------
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

def parse_resume_with_llm(resume_text: str, template_structure: Dict, llm_client: ZhipuAIClient) -> Dict:
    """
    让AI根据模板结构输出JSON，包含单字段和表格数据。
    输出示例：
    {
        "fields": {"姓名": "张三", "电话": "138...", ...},
        "工作经历": [{"工作开始日期": "2020-01", "单位名称": "XX公司", ...}, ...],
        "项目经历": [...]
    }
    """
    # 构建描述
    field_labels = list(template_structure['key_value_fields'].keys())
    table_descs = []
    for tbl in template_structure['tables']:
        table_descs.append(f"- 表格「{tbl['name']}」列头：{tbl['headers']}")
    
    prompt = f"""
你是一个专业的简历解析助手。请根据以下简历文本，提取信息，并以严格的JSON格式输出。

需要提取的信息分为两部分：

1. 单个字段（键值对）：
{field_labels}

2. 多个表格（每个表格是一个对象数组）：
{chr(10).join(table_descs) if table_descs else '无'}

输出格式（严格JSON，不要解释）：
{{
    "fields": {{"字段1": "值1", "字段2": "值2", ...}},
    "工作经历": [{{"列头1": "值1", "列头2": "值2", ...}}, ...],
    "项目经历": [{{...}}]
}}

注意：
- 字段名必须与给定的完全一致。
- 对于工作经历和项目经历，请按时间由近及远排序，提取所有相关条目。
- 如果某个字段找不到信息，填写空字符串""。
- 如果某个表格没有数据，输出空数组[]。

简历文本：
{resume_text}
"""
    response = llm_client.call(prompt)
    # 清理markdown
    if response.startswith("```json"):
        response = response[7:]
    if response.endswith("```"):
        response = response[:-3]
    data = json.loads(response.strip())
    # 确保结构完整
    if 'fields' not in data:
        data['fields'] = {}
    for tbl in [t['name'] for t in template_structure['tables']]:
        if tbl not in data:
            data[tbl] = []
    return data

def fill_template_keep_format(template_path: str, parsed_data: Dict, output_path: str, template_structure: Dict):
    """根据扫描的结构填充Excel，保留原格式"""
    wb = openpyxl.load_workbook(template_path)
    ws = wb[template_structure['sheet_name']]
    
    # 1. 填充键值对字段
    fields = parsed_data.get('fields', {})
    for label, (row, col) in template_structure['key_value_fields'].items():
        if label in fields:
            ws.cell(row=row, column=col, value=fields[label])
    
    # 2. 填充表格数据
    for tbl in template_structure['tables']:
        table_name = tbl['name']
        data_rows = parsed_data.get(table_name, [])
        if not data_rows:
            continue
        # 确定表格数据区域：从 data_start_row 开始，先删除原有数据行
        start_row = tbl['data_start_row']
        # 找到表格结束行（连续非空行的最大行，或遇到下一个表格/空行）
        end_row = start_row
        while end_row <= ws.max_row:
            # 检查该行是否有任何非空单元格（在表头列范围内）
            has_data = False
            for col in range(tbl['start_col'], tbl['start_col'] + len(tbl['headers'])):
                if ws.cell(end_row, col).value:
                    has_data = True
                    break
            if not has_data:
                break
            end_row += 1
        # 删除原有数据行
        if end_row > start_row:
            ws.delete_rows(start_row, end_row - start_row)
        # 写入新数据
        for i, record in enumerate(data_rows):
            current_row = start_row + i
            for col_idx, header in enumerate(tbl['headers']):
                if header in record:
                    value = record[header]
                    ws.cell(row=current_row, column=tbl['start_col'] + col_idx, value=value)
    wb.save(output_path)

# ----------------------------- Streamlit界面 -----------------------------
with st.sidebar:
    st.header("🔑 智谱AI 配置")
    api_key = st.text_input("API Key", type="password", help="从 https://bigmodel.cn 获取")
    model_name = st.selectbox("模型", ["glm-4-flash", "glm-4-plus"], index=0)
    st.markdown("---")
    st.markdown("**模板要求**：")
    st.markdown("- 键值对字段以冒号结尾（如 `姓名：`）")
    st.markdown("- 表格需有明确的表头行（如 `工作开始日期`、`单位名称`）")
    st.markdown("- 支持合并单元格")

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
    # 1. 扫描模板结构
    with st.spinner("解析模板结构..."):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp.write(template_file.getbuffer())
            tmp_path = tmp.name
        try:
            structure = scan_template_structure(tmp_path)
            st.success("模板结构识别成功")
            with st.expander("查看识别的字段和表格"):
                st.write("键值对字段：", list(structure['key_value_fields'].keys()))
                for t in structure['tables']:
                    st.write(f"表格 {t['name']} 列头：", t['headers'])
        except Exception as e:
            st.error(f"模板解析失败：{e}")
            st.stop()
    progress.progress(20)
    
    # 2. 提取PDF
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
    
    # 3. AI解析
    with st.spinner(f"调用智谱AI ({model_name}) 解析..."):
        try:
            client = ZhipuAIClient(api_key, model_name)
            parsed_data = parse_resume_with_llm(resume_text, structure, client)
            st.success("AI解析完成")
            with st.expander("查看提取的数据"):
                st.json(parsed_data)
        except Exception as e:
            st.error(f"AI解析失败：{e}")
            st.stop()
    progress.progress(70)
    
    # 4. 填充并导出
    with st.spinner("生成Excel（保留格式）..."):
        try:
            output_temp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx").name
            fill_template_keep_format(tmp_path, parsed_data, output_temp, structure)
            with open(output_temp, "rb") as f:
                excel_data = f.read()
            st.download_button("📥 下载生成的Excel", data=excel_data, file_name="filled_resume.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            os.unlink(tmp_path)
            os.unlink(output_temp)
        except Exception as e:
            st.error(f"生成失败：{e}")
            st.stop()
    progress.progress(100)
    st.balloons()
