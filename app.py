import streamlit as st
import tempfile
import os
import json
import pdfplumber
import pandas as pd
from openai import OpenAI
from typing import Dict, Any

# ----------------------------- 页面配置 -----------------------------
st.set_page_config(page_title="AI 简历解析与填充工具", layout="wide")
st.title("📄 AI 简历解析与 Excel 模板填充")
st.markdown("上传 PDF 简历和 Excel 模板，AI 自动提取信息并按模板生成完整简历。")

# ----------------------------- 辅助函数 -----------------------------
def extract_text_from_pdf(pdf_file) -> str:
    """从上传的 PDF 文件对象中提取文本"""
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

def parse_resume_with_llm(resume_text: str, api_key: str, model: str) -> Dict[str, Any]:
    """调用 OpenAI 解析简历，返回结构化 JSON"""
    client = OpenAI(api_key=api_key)
    prompt = f"""
你是一个简历解析专家。请根据以下简历文本，提取关键信息，并以严格的 JSON 格式输出。
输出必须包含以下字段：
- name (字符串)
- phone (字符串)
- email (字符串)
- work_experiences (列表，每个元素包含 company, position, duration, description)
- education_experiences (列表，每个元素包含 school, degree, duration)

如果某个字段没有找到，请用空字符串或空列表表示。
只输出 JSON，不要有其他解释。

简历文本：
{resume_text}
"""
    response = client.chat.completions.create(
        model=model,
        messages=[{"role": "user", "content": prompt}],
        temperature=0.0
    )
    content = response.choices[0].message.content
    # 清理可能的 markdown 标记
    if content.startswith("```json"):
        content = content[7:]
    if content.endswith("```"):
        content = content[:-3]
    return json.loads(content.strip())

def fill_excel_from_template(template_file, parsed_data: Dict[str, Any], output_path: str):
    """
    根据模板 Excel 填充数据。
    假设模板包含三个工作表：基本信息、工作经历、教育经历
    列头分别为：
       基本信息：姓名,电话,邮箱
       工作经历：公司名称,职位,工作时间,工作描述
       教育经历：学校,学位,时间
    """
    excel_file = pd.ExcelFile(template_file)
    sheet_names = excel_file.sheet_names

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        for sheet_name in sheet_names:
            df = pd.read_excel(template_file, sheet_name=sheet_name, header=0)
            if sheet_name == "基本信息":
                if "姓名" in df.columns:
                    df.at[0, "姓名"] = parsed_data.get("name", "")
                if "电话" in df.columns:
                    df.at[0, "电话"] = parsed_data.get("phone", "")
                if "邮箱" in df.columns:
                    df.at[0, "邮箱"] = parsed_data.get("email", "")
            elif sheet_name == "工作经历":
                work_list = parsed_data.get("work_experiences", [])
                if work_list:
                    work_df = pd.DataFrame(work_list)
                    work_df.rename(columns={
                        "company": "公司名称",
                        "position": "职位",
                        "duration": "工作时间",
                        "description": "工作描述"
                    }, inplace=True)
                    for col in df.columns:
                        if col not in work_df.columns:
                            work_df[col] = ""
                    df = pd.concat([df, work_df], ignore_index=True)
            elif sheet_name == "教育经历":
                edu_list = parsed_data.get("education_experiences", [])
                if edu_list:
                    edu_df = pd.DataFrame(edu_list)
                    edu_df.rename(columns={
                        "school": "学校",
                        "degree": "学位",
                        "duration": "时间"
                    }, inplace=True)
                    for col in df.columns:
                        if col not in edu_df.columns:
                            edu_df[col] = ""
                    df = pd.concat([df, edu_df], ignore_index=True)
            df.to_excel(writer, sheet_name=sheet_name, index=False)

# ----------------------------- UI 布局 -----------------------------
with st.sidebar:
    st.header("⚙️ 配置")
    api_key = st.text_input("OpenAI API Key", type="password", help="输入你的 OpenAI API Key，或设置环境变量 OPENAI_API_KEY")
    model = st.selectbox("选择模型", ["gpt-3.5-turbo", "gpt-4"], index=0)
    st.markdown("---")
    st.markdown("**模板要求**")
    st.markdown("""
    - 工作表 `基本信息` 需包含列：`姓名`、`电话`、`邮箱`
    - 工作表 `工作经历` 需包含列：`公司名称`、`职位`、`工作时间`、`工作描述`
    - 工作表 `教育经历` 需包含列：`学校`、`学位`、`时间`
    """)

col1, col2 = st.columns(2)
with col1:
    pdf_file = st.file_uploader("上传 PDF 简历", type=["pdf"])
with col2:
    template_file = st.file_uploader("上传 Excel 模板", type=["xlsx"])

if st.button("🚀 开始解析并生成简历", type="primary"):
    if not pdf_file or not template_file:
        st.error("请同时上传 PDF 简历和 Excel 模板。")
        st.stop()
    if not api_key:
        st.error("请输入 OpenAI API Key。")
        st.stop()

    # 1. 提取 PDF 文本
    with st.spinner("正在读取 PDF 简历..."):
        try:
            resume_text = extract_text_from_pdf(pdf_file)
            if not resume_text:
                st.error("PDF 内容为空或无法解析。")
                st.stop()
        except Exception as e:
            st.error(f"读取 PDF 失败: {e}")
            st.stop()

    # 2. 调用大模型
    with st.spinner("AI 正在解析简历信息（可能需要几秒）..."):
        try:
            parsed_data = parse_resume_with_llm(resume_text, api_key, model)
            st.success("解析完成！")
            with st.expander("查看 AI 提取的 JSON 数据"):
                st.json(parsed_data)
        except Exception as e:
            st.error(f"AI 解析失败: {e}")
            st.stop()

    # 3. 填充模板并生成 Excel
    with st.spinner("正在生成 Excel 文件..."):
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_output:
                output_path = tmp_output.name
            fill_excel_from_template(template_file, parsed_data, output_path)
            # 提供下载按钮
            with open(output_path, "rb") as f:
                excel_bytes = f.read()
            st.download_button(
                label="📥 下载生成的 Excel 简历",
                data=excel_bytes,
                file_name="filled_resume.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            os.unlink(output_path)
        except Exception as e:
            st.error(f"生成 Excel 失败: {e}")
            st.stop()

    st.balloons()
