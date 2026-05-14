import streamlit as st
import tempfile
import os
import json
import pdfplumber
import openpyxl
from openpyxl.utils import get_column_letter
from typing import Dict, Any, List, Union
import pandas as pd
from openai import OpenAI

# ----------------------------- 页面配置 -----------------------------
st.set_page_config(page_title="AI 简历解析与模板填充 (智谱AI版)", layout="wide")
st.title("📄 AI 简历解析 → 任意 Excel 模板 (仅使用第一个Sheet)")
st.markdown("上传 PDF 简历和 Excel 模板，**智谱AI** 自动理解模板结构并填充内容，**保留原格式**。")

# ----------------------------- 1. 模板结构提取 (仅第一个Sheet) -----------------------------
def get_template_structure(template_path: str) -> Dict[str, List[str]]:
    """
    只读取 Excel 模板的第一个 Sheet，返回 {sheet_name: [列头列表]}
    """
    wb = openpyxl.load_workbook(template_path, data_only=True)
    sheet_name = wb.sheetnames[0]  # 只取第一个 sheet
    ws = wb[sheet_name]
    headers = []
    for col in range(1, ws.max_column + 1):
        cell_value = ws.cell(row=1, column=col).value
        if cell_value:
            headers.append(str(cell_value).strip())
    return {sheet_name: headers}

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
    """智谱AI客户端（兼容OpenAI SDK）"""
    def __init__(self, api_key: str, model_name: str = "glm-4-flash"):
        self.api_key = api_key
        self.model_name = model_name
        self.client = OpenAI(
            api_key=api_key,
            base_url="https://open.bigmodel.cn/api/paas/v4/"
        )
    
    def call(self, prompt: str) -> str:
        response = self.client.chat.completions.create(
            model=self.model_name,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.0
        )
        return response.choices[0].message.content

# ----------------------------- 4. 智能解析（根据模板结构） -----------------------------
def parse_resume_with_llm(
    resume_text: str,
    template_structure: Dict[str, List[str]],
    llm_client: ZhipuAIClient
) -> Dict[str, Union[Dict, List]]:
    """
    让 AI 根据模板结构返回 JSON。
    每个 sheet 对应一个 JSON 对象，如果是多条数据则为数组。
    """
    # 构建描述模板结构的 prompt
    structure_desc = ""
    for sheet, headers in template_structure.items():
        structure_desc += f"- Sheet「{sheet}」列头：{headers}\n"
    prompt = f"""
你是一个专业的简历解析助手。请根据下方简历文本，按照给定的 Excel 模板结构输出 JSON 数据。

模板结构：
{structure_desc}

要求：
1. 对于每个 Sheet，根据列头提取相应信息。
2. 如果某个 Sheet 预期只有一行数据（例如“基本信息”），输出一个对象（dict）。
3. 如果某个 Sheet 预期有多行数据（例如“工作经历”、“教育经历”），输出一个对象数组（list of dict）。
4. 输出必须是严格的 JSON 格式，不要有任何额外解释或 markdown 标记。
5. 如果某个字段找不到信息，填写空字符串 ""。

简历文本：
{resume_text}
"""
    response = llm_client.call(prompt)
    # 清理可能的 markdown
    if response.startswith("```json"):
        response = response[7:]
    if response.endswith("```"):
        response = response[:-3]
    try:
        data = json.loads(response.strip())
    except json.JSONDecodeError as e:
        st.error(f"AI 返回的不是合法 JSON：{response[:200]}")
        raise e
    return data

# ----------------------------- 5. 通用填充（保持格式，仅第一个Sheet） -----------------------------
def fill_template_keep_format(
    template_path: str,
    parsed_data: Dict[str, Union[Dict, List]],
    output_path: str
):
    """
    用 openpyxl 填充数据，保留原始样式、合并单元格。只处理模板的第一个 sheet。
    """
    wb = openpyxl.load_workbook(template_path)
    target_sheet_name = wb.sheetnames[0]  # 只处理第一个 sheet
    
    # 获取 AI 返回的对应 sheet 数据，若不存在则用空数据
    if target_sheet_name not in parsed_data:
        st.warning(f"AI 返回的数据中没有找到 sheet '{target_sheet_name}'，将使用空数据填充。")
        data = {}
    else:
        data = parsed_data[target_sheet_name]
    
    ws = wb[target_sheet_name]
    # 获取第一行表头（非空单元格值）
    headers = []
    for cell in ws[1]:
        if cell.value:
            headers.append(str(cell.value).strip())
    # 建立列头到列索引的映射
    col_index = {}
    for idx, h in enumerate(headers, start=1):
        col_index[h] = idx

    # 处理单行数据（dict）
    if isinstance(data, dict):
        row_num = 2
        for field, value in data.items():
            if field in col_index:
                ws.cell(row=row_num, column=col_index[field], value=value)
    # 处理多行数据（list）
    elif isinstance(data, list):
        # 删除第2行及以下所有数据行（保留表头）
        if ws.max_row >= 2:
            ws.delete_rows(2, ws.max_row - 1)
        # 写入新数据
        for row_idx, record in enumerate(data, start=2):
            for field, value in record.items():
                if field in col_index:
                    ws.cell(row=row_idx, column=col_index[field], value=value)
    else:
        st.warning(f"Sheet {target_sheet_name} 数据格式异常：{type(data)}，跳过")
    
    wb.save(output_path)

# ----------------------------- 6. Streamlit 界面（仅智谱AI） -----------------------------
with st.sidebar:
    st.header("🔑 智谱AI 配置")
    api_key = st.text_input("API Key", type="password", help="从 https://bigmodel.cn 获取，新用户送免费额度")
    model_name = st.selectbox(
        "模型选择",
        ["glm-4-flash", "glm-4-plus", "glm-4-air", "glm-4-long"],
        index=0,
        help="glm-4-flash 永久免费，推荐使用"
    )
    st.markdown("---")
    st.markdown("**说明**：")
    st.markdown("- 上传任意 Excel 模板，AI 会根据**第一个Sheet的列头**自动填充")
    st.markdown("- 保留原有样式、合并单元格")
    st.markdown("- 支持多行数据（如工作经历、教育经历）")
    st.markdown("- 智谱AI新用户赠送免费额度，glm-4-flash 长期免费")

col1, col2 = st.columns(2)
with col1:
    pdf_file = st.file_uploader("📄 上传 PDF 简历", type=["pdf"])
with col2:
    template_file = st.file_uploader("📊 上传 Excel 模板", type=["xlsx"])

if st.button("🚀 开始解析并生成", type="primary"):
    if not pdf_file or not template_file:
        st.error("请同时上传 PDF 和 Excel 模板。")
        st.stop()
    if not api_key:
        st.error("请填写智谱AI API Key。")
        st.stop()

    progress_bar = st.progress(0, text="准备就绪...")

    # 1. 提取模板结构（仅第一个Sheet）
    with st.spinner("读取模板结构..."):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_template:
            tmp_template.write(template_file.getbuffer())
            tmp_template_path = tmp_template.name
        try:
            template_structure = get_template_structure(tmp_template_path)
            st.success("模板结构读取成功")
            with st.expander("查看模板结构（仅第一个Sheet）"):
                st.json(template_structure)
        except Exception as e:
            st.error(f"模板解析失败：{e}")
            st.stop()
    progress_bar.progress(20, text="模板结构已读取")

    # 2. 提取 PDF 文本
    with st.spinner("读取 PDF 简历..."):
        try:
            resume_text = extract_text_from_pdf(pdf_file)
            if not resume_text:
                st.error("PDF 内容为空或无法解析。")
                st.stop()
        except Exception as e:
            st.error(f"PDF 读取失败：{e}")
            st.stop()
    progress_bar.progress(40, text="PDF 内容提取完成")

    # 3. 调用智谱AI解析
    with st.spinner(f"正在调用智谱AI（{model_name}）解析简历..."):
        try:
            llm_client = ZhipuAIClient(api_key=api_key, model_name=model_name)
            parsed_data = parse_resume_with_llm(resume_text, template_structure, llm_client)
            st.success("AI 解析完成！")
            with st.expander("查看 AI 提取的数据"):
                st.json(parsed_data)
        except Exception as e:
            st.error(f"AI 解析失败：{e}")
            st.stop()
    progress_bar.progress(70, text="AI 解析完成")

    # 4. 填充模板并输出（仅第一个Sheet）
    with st.spinner("生成 Excel 文件（保留原格式）..."):
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_output:
                output_path = tmp_output.name
            fill_template_keep_format(tmp_template_path, parsed_data, output_path)
            # 提供下载
            with open(output_path, "rb") as f:
                excel_bytes = f.read()
            st.download_button(
                label="📥 下载生成的 Excel 简历",
                data=excel_bytes,
                file_name="filled_resume.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            os.unlink(output_path)
            os.unlink(tmp_template_path)
        except Exception as e:
            st.error(f"生成 Excel 失败：{e}")
            st.stop()
    progress_bar.progress(100, text="完成！")
    st.balloons()
