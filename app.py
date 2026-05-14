import streamlit as st
import tempfile
import os
import json
import requests
import pdfplumber
import openpyxl
from openpyxl.utils import get_column_letter
from typing import Dict, Any, List, Union
import pandas as pd  # 仅用于读取列头，不用于写入

# ----------------------------- 页面配置 -----------------------------
st.set_page_config(page_title="AI 简历解析与模板填充 (通用版)", layout="wide")
st.title("📄 AI 简历解析 → 任意 Excel 模板")
st.markdown("上传 PDF 简历和 Excel 模板，AI 自动理解模板结构并填充内容，**保留原格式**。")

# ----------------------------- 1. 模板结构提取 -----------------------------
def get_template_structure(template_path: str) -> Dict[str, List[str]]:
    """
    读取 Excel 模板的所有 Sheet，返回 {sheet_name: [列头列表]}
    使用 openpyxl 读取第一行作为列头。
    """
    wb = openpyxl.load_workbook(template_path, data_only=True)
    structure = {}
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        headers = []
        for col in range(1, ws.max_column + 1):
            cell_value = ws.cell(row=1, column=col).value
            if cell_value:
                headers.append(str(cell_value).strip())
        structure[sheet_name] = headers
    return structure

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

# ----------------------------- 3. 大模型调用抽象层 -----------------------------
class ModelClient:
    """统一接口，支持 OpenAI 和 Ollama"""
    def __init__(self, model_type: str, api_key: str = None, model_name: str = None, base_url: str = None):
        self.model_type = model_type
        self.api_key = api_key
        self.model_name = model_name
        self.base_url = base_url

    def call(self, prompt: str) -> str:
        if self.model_type == "OpenAI":
            from openai import OpenAI
            client = OpenAI(api_key=self.api_key)
            response = client.chat.completions.create(
                model=self.model_name,
                messages=[{"role": "user", "content": prompt}],
                temperature=0.0
            )
            return response.choices[0].message.content
        elif self.model_type == "Ollama":
            url = f"{self.base_url}/api/generate"
            payload = {
                "model": self.model_name,
                "prompt": prompt,
                "stream": False,
                "options": {"temperature": 0.0}
            }
            resp = requests.post(url, json=payload)
            resp.raise_for_status()
            return resp.json()["response"]
        else:
            raise ValueError(f"不支持的模型类型: {self.model_type}")

# ----------------------------- 4. 智能解析（根据模板结构） -----------------------------
def parse_resume_with_llm(
    resume_text: str,
    template_structure: Dict[str, List[str]],
    model_client: ModelClient
) -> Dict[str, Union[Dict, List]]:
    """
    让 AI 根据模板结构返回 JSON。
    每个 sheet 对应一个 JSON 对象，如果是多条数据则为数组。
    输出示例：
    {
        "基本信息": {"姓名": "张三", "电话": "138...", ...},
        "工作经历": [{"公司名称": "XX", "职位": "工程师", ...}, ...],
        "教育经历": [{"学校": "XX大学", ...}]
    }
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
    response = model_client.call(prompt)
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

# ----------------------------- 5. 通用填充（保持格式，使用 openpyxl） -----------------------------
def fill_template_keep_format(
    template_path: str,
    parsed_data: Dict[str, Union[Dict, List]],
    output_path: str
):
    """
    用 openpyxl 填充数据，保留原始样式、合并单元格。
    策略：
      - 保留第一行表头不变。
      - 对于单行数据（Dict）：覆盖第二行（如果第二行有数据则替换，无则创建）。
      - 对于多行数据（List）：从第二行开始，先清空原有数据行（保留表头），然后逐行写入列表数据。
    """
    wb = openpyxl.load_workbook(template_path)
    for sheet_name, data in parsed_data.items():
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        headers = [cell.value for cell in ws[1] if cell.value]  # 第一行列头
        # 建立列头到列索引的映射
        col_index = {}
        for idx, h in enumerate(headers, start=1):
            col_index[str(h).strip()] = idx

        # 处理单行数据（dict）
        if isinstance(data, dict):
            # 写入第二行
            row_num = 2
            for field, value in data.items():
                if field in col_index:
                    ws.cell(row=row_num, column=col_index[field], value=value)
        # 处理多行数据（list）
        elif isinstance(data, list):
            # 先删除第2行及以下的所有数据行（保留表头）
            if ws.max_row >= 2:
                ws.delete_rows(2, ws.max_row - 1)
            # 写入新数据
            for row_idx, record in enumerate(data, start=2):
                for field, value in record.items():
                    if field in col_index:
                        ws.cell(row=row_idx, column=col_index[field], value=value)
        else:
            st.warning(f"Sheet {sheet_name} 数据格式异常：{type(data)}，跳过")
    wb.save(output_path)

# ----------------------------- 6. Streamlit 界面 -----------------------------
with st.sidebar:
    st.header("🤖 模型配置")
    model_type = st.selectbox("选择大模型后端", ["OpenAI", "Ollama"])
    if model_type == "OpenAI":
        api_key = st.text_input("OpenAI API Key", type="password", help="需付费，但效果好")
        model_name = st.selectbox("模型", ["gpt-3.5-turbo", "gpt-4"])
        base_url = None
    else:  # Ollama
        api_key = None
        model_name = st.text_input("Ollama 模型名称", value="llama3", help="本地已安装的模型，如 llama3, qwen2")
        base_url = st.text_input("Ollama API 地址", value="http://localhost:11434")
    st.markdown("---")
    st.markdown("**说明**：")
    st.markdown("- 上传任意 Excel 模板，AI 会根据列头自动填充")
    st.markdown("- 保留原有样式、合并单元格")
    st.markdown("- 支持多行数据（如工作经历、教育经历）")

col1, col2 = st.columns(2)
with col1:
    pdf_file = st.file_uploader("📄 上传 PDF 简历", type=["pdf"])
with col2:
    template_file = st.file_uploader("📊 上传 Excel 模板", type=["xlsx"])

if st.button("🚀 开始解析并生成", type="primary"):
    if not pdf_file or not template_file:
        st.error("请同时上传 PDF 和 Excel 模板。")
        st.stop()
    if model_type == "OpenAI" and not api_key:
        st.error("请填写 OpenAI API Key。")
        st.stop()
    if model_type == "Ollama" and not model_name:
        st.error("请填写 Ollama 模型名称。")
        st.stop()

    progress_bar = st.progress(0, text="准备就绪...")

    # 1. 提取模板结构
    with st.spinner("读取模板结构..."):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_template:
            tmp_template.write(template_file.getbuffer())
            tmp_template_path = tmp_template.name
        try:
            template_structure = get_template_structure(tmp_template_path)
            st.success("模板结构读取成功")
            with st.expander("查看模板结构"):
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

    # 3. 调用 AI 解析
    with st.spinner(f"正在调用 {model_type} 解析简历（可能需要几秒）..."):
        try:
            model_client = ModelClient(model_type, api_key, model_name, base_url)
            parsed_data = parse_resume_with_llm(resume_text, template_structure, model_client)
            st.success("AI 解析完成！")
            with st.expander("查看 AI 提取的数据"):
                st.json(parsed_data)
        except Exception as e:
            st.error(f"AI 解析失败：{e}")
            st.stop()
    progress_bar.progress(70, text="AI 解析完成")

    # 4. 填充模板并输出
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
