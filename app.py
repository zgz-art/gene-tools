import streamlit as st
import os
import tempfile
import json
from datetime import datetime
import re
import pypdf
from zhipuai import ZhipuAI
import openpyxl
from openpyxl.styles import numbers
from copy import copy

# ==================== 页面配置 ====================
st.set_page_config(page_title="简历智能填充工具", page_icon="📄", layout="wide")

# 自定义CSS
st.markdown("""
<style>
    .reportview-container .main .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
    }
    .stButton button {
        background-color: #4CAF50;
        color: white;
        border-radius: 8px;
        padding: 0.5rem 1rem;
        font-weight: bold;
    }
    .css-1v3fvcr {
        background-color: #f8f9fa;
        border-radius: 12px;
        padding: 20px;
    }
    h1 {
        color: #2c3e50;
        text-align: center;
    }
</style>
""", unsafe_allow_html=True)

# ==================== 辅助函数 ====================
def normalize_date(date_str: str) -> str:
    """将各种日期格式转为 YYYY-MM-DD，缺失日则补当月首日"""
    if not date_str:
        return ""
    date_str = str(date_str).strip()
    for fmt in ["%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y年%m月%d日", "%Y-%m", "%Y/%m", "%Y.%m", "%Y年%m月"]:
        try:
            if fmt.endswith("%m"):
                dt = datetime.strptime(date_str, fmt)
                return dt.strftime("%Y-%m-01")
            else:
                dt = datetime.strptime(date_str, fmt)
                return dt.strftime("%Y-%m-%d")
        except ValueError:
            continue
    year_match = re.search(r"\b(19|20)\d{2}\b", date_str)
    if year_match:
        return f"{year_match.group()}-01-01"
    return date_str

def extract_text_from_pdf(pdf_file) -> str:
    """从上传的 PDF 文件中提取纯文本"""
    reader = pypdf.PdfReader(pdf_file)
    text = ""
    for page in reader.pages:
        text += page.extract_text() + "\n"
    return text

# ==================== AI 提取函数 ====================
SYSTEM_PROMPT = """
你是一个专业的简历信息提取助手。请从以下简历文本中提取指定字段，并以 JSON 格式返回。
要求：
1. 日期统一使用 "YYYY-MM-DD" 格式，若只提供年月，则默认补充 "-01"。
2. 若某字段不存在，返回空字符串或空列表。
3. 工作经历和项目经历请按时间由近及远排序（开始日期越晚越靠前）。

输出 JSON 结构如下：
{
    "basic": {
        "姓名": "",
        "身份证号": "",
        "出生日期": "",
        "电话": "",
        "首次参加工作时间": "",
        "首次参加IT领域工作时间": "",
        "最高学历": "",
        "掌握语言": "",
        "掌握技能": "",
        "专业证书": ""
    },
    "education": {
        "undergraduate": {
            "入学时间": "",
            "毕业院校": "",
            "毕业时间": "",
            "专业": "",
            "毕业证编号": "",
            "毕业证学信网在线验证码": "",
            "学位证编号": "",
            "学位证学信网在线验证码": ""
        },
        "postgraduate": {
            "入学时间": "",
            "毕业院校": "",
            "毕业时间": "",
            "专业": "",
            "毕业证编号": "",
            "毕业证学信网在线验证码": "",
            "学位证编号": "",
            "学位证学信网在线验证码": ""
        }
    },
    "work_experience": [
        {
            "开始日期": "",
            "结束日期": "",
            "单位名称": "",
            "岗位": ""
        }
    ],
    "project_experience": [
        {
            "开始日期": "",
            "结束日期": "",
            "项目名称": "",
            "项目描述": "",
            "项目角色": ""
        }
    ],
    "extra": {
        "供应商缩写": "",
        "类型": "",
        "岗位": "",
        "级别": ""
    }
}
"""

def extract_resume_info(api_key: str, pdf_text: str, model: str = "glm-4-plus") -> dict:
    """调用智谱 AI 提取信息，支持指定模型"""
    client = ZhipuAI(api_key=api_key)
    response = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": f"简历文本：\n{pdf_text}"}
        ],
        temperature=0.1,
        response_format={"type": "json_object"}
    )
    result_text = response.choices[0].message.content
    data = json.loads(result_text)
    return data

# ==================== Excel 填充函数 ====================
def find_cell_by_value(worksheet, value, exact=False):
    """遍历查找单元格（支持模糊匹配或精确匹配），返回 (row, col) 或 None"""
    for row in worksheet.iter_rows():
        for cell in row:
            cell_val = cell.value
            if cell_val is None:
                continue
            if exact:
                if cell_val == value:
                    return cell.row, cell.column
            else:
                if value.lower() in str(cell_val).lower():
                    return cell.row, cell.column
    return None

def find_row_with_keyword(worksheet, keyword):
    """返回包含关键字的行索引（行号从1开始）"""
    for row in range(1, worksheet.max_row + 1):
        for col in range(1, worksheet.max_column + 1):
            val = worksheet.cell(row, col).value
            if val and keyword in str(val):
                return row
    return None

def copy_row_style(source_row, target_row, worksheet, max_col):
    """复制源行样式到目标行（用于新增行时保留格式）"""
    for col in range(1, max_col + 1):
        source_cell = worksheet.cell(source_row, col)
        target_cell = worksheet.cell(target_row, col)
        if source_cell.has_style:
            target_cell.font = copy(source_cell.font)
            target_cell.border = copy(source_cell.border)
            target_cell.fill = copy(source_cell.fill)
            target_cell.number_format = copy(source_cell.number_format)
            target_cell.alignment = copy(source_cell.alignment)

def insert_rows(worksheet, start_row, count, max_col):
    """在 start_row 之前插入 count 行，并复制 start_row 的样式"""
    worksheet.insert_rows(start_row, count)
    for i in range(count):
        new_row = start_row + i
        copy_row_style(start_row + count, new_row, worksheet, max_col)

def clear_rows(worksheet, start_row, end_row, max_col):
    """清空指定行区域的内容（保留样式）"""
    for row in range(start_row, end_row + 1):
        for col in range(1, max_col + 1):
            worksheet.cell(row, col).value = None

def fill_basic_info(worksheet, basic_data):
    """模糊匹配基础信息字段，填充到右侧单元格（假设字段在A列，右侧B列）"""
    for key, value in basic_data.items():
        if not value:
            continue
        pos = find_cell_by_value(worksheet, key, exact=False)
        if pos:
            row, col = pos
            target_cell = worksheet.cell(row, col + 1)
            target_cell.value = value

def fill_education_block(worksheet, keyword, edu_data):
    """
    精确匹配 keyword（如“本科学历”）所在行，然后处理接下来的若干行。
    假设格式：A列为字段名，B列为待填充值。
    edu_data 为字典，字段名与模板中的文本进行精确匹配（忽略空格）。
    """
    pos = find_cell_by_value(worksheet, keyword, exact=True)
    if not pos:
        return
    row_start = pos[0] + 1
    for offset in range(10):
        current_row = row_start + offset
        field_cell = worksheet.cell(current_row, 1)
        if field_cell.value is None:
            break
        field_name = str(field_cell.value).strip()
        for data_key, data_val in edu_data.items():
            if data_val and data_key == field_name:
                target_cell = worksheet.cell(current_row, 2)
                target_cell.value = data_val
                break

def fill_work_experience(worksheet, work_list):
    keyword_row = find_row_with_keyword(worksheet, "工作经历（由近及远，仅限IT相关经历）")
    if not keyword_row:
        return
    header_row = keyword_row + 1
    data_start_row = header_row + 1
    reserved_rows = 3
    current_data_end = data_start_row + reserved_rows - 1
    max_col = worksheet.max_column
    clear_rows(worksheet, data_start_row, current_data_end, max_col)
    
    sorted_work = sorted(work_list, key=lambda x: x.get("开始日期", "1900-01-01"), reverse=True)
    need_rows = len(sorted_work)
    if need_rows > reserved_rows:
        insert_rows(worksheet, current_data_end + 1, need_rows - reserved_rows, max_col)
    
    for idx, work in enumerate(sorted_work):
        target_row = data_start_row + idx
        for col in range(1, max_col + 1):
            header_val = worksheet.cell(header_row, col).value
            if not header_val:
                continue
            if "开始日期" in str(header_val):
                worksheet.cell(target_row, col).value = normalize_date(work.get("开始日期", ""))
            elif "结束日期" in str(header_val):
                worksheet.cell(target_row, col).value = normalize_date(work.get("结束日期", ""))
            elif "单位名称" in str(header_val):
                worksheet.cell(target_row, col).value = work.get("单位名称", "")
            elif "岗位" in str(header_val):
                worksheet.cell(target_row, col).value = work.get("岗位", "")

def fill_project_experience(worksheet, project_list):
    keyword_row = find_row_with_keyword(worksheet, "项目经历（与上述工作经历匹配，仅IT相关经历）")
    if not keyword_row:
        return
    header_row = keyword_row + 1
    data_start_row = header_row + 1
    reserved_rows = 6
    current_data_end = data_start_row + reserved_rows - 1
    max_col = worksheet.max_column
    clear_rows(worksheet, data_start_row, current_data_end, max_col)
    
    sorted_proj = sorted(project_list, key=lambda x: x.get("开始日期", "1900-01-01"), reverse=True)
    need_rows = len(sorted_proj)
    if need_rows > reserved_rows:
        insert_rows(worksheet, current_data_end + 1, need_rows - reserved_rows, max_col)
    
    for idx, proj in enumerate(sorted_proj):
        target_row = data_start_row + idx
        for col in range(1, max_col + 1):
            header_val = worksheet.cell(header_row, col).value
            if not header_val:
                continue
            if "开始日期" in str(header_val):
                worksheet.cell(target_row, col).value = normalize_date(proj.get("开始日期", ""))
            elif "结束日期" in str(header_val):
                worksheet.cell(target_row, col).value = normalize_date(proj.get("结束日期", ""))
            elif "项目名称" in str(header_val):
                worksheet.cell(target_row, col).value = proj.get("项目名称", "")
            elif "项目描述" in str(header_val):
                worksheet.cell(target_row, col).value = proj.get("项目描述", "")
            elif "项目角色" in str(header_val):
                worksheet.cell(target_row, col).value = proj.get("项目角色", "")

def fill_template(template_path, output_path, ai_data):
    """主填充函数"""
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active
    
    fill_basic_info(ws, ai_data.get("basic", {}))
    
    under = ai_data.get("education", {}).get("undergraduate", {})
    if under:
        fill_education_block(ws, "本科学历", under)
    
    post = ai_data.get("education", {}).get("postgraduate", {})
    if post:
        fill_education_block(ws, "研究生学历", post)
    
    work_list = ai_data.get("work_experience", [])
    if work_list:
        fill_work_experience(ws, work_list)
    
    proj_list = ai_data.get("project_experience", [])
    if proj_list:
        fill_project_experience(ws, proj_list)
    
    wb.save(output_path)

# ==================== Streamlit UI ====================
st.title("📄 智能简历填充工具")
st.markdown("上传您的原始简历（PDF）和简历模板（Excel），AI 将自动提取信息并生成标准格式文件。")

# 侧边栏配置
with st.sidebar:
    st.header("⚙️ 配置")
    api_key = st.text_input("🔑 智谱 AI API Key", type="password", help="前往 https://open.bigmodel.cn/ 获取")
    model_name = st.selectbox(
        "🤖 选择大模型",
        options=["glm-4-plus", "glm-4-flash", "glm-4-air", "glm-4-long"],
        index=0,
        help="不同模型的速度、成本与效果有所差异，推荐使用 glm-4-plus 获得最佳提取质量"
    )
    st.markdown("---")
    st.caption("需要帮助？请查看 [使用说明](https://github.com/your-repo)")

with st.expander("📌 使用说明", expanded=True):
    st.markdown("""
    1. 准备 **PDF 格式** 的原始简历文件  
    2. 准备 **Excel 模板**（.xlsx），模板中应包含以下关键字：  
       - 基础信息字段（姓名、身份证号...）  
       - 精确的 **“本科学历”**、**“研究生学历”** 文字  
       - **“工作经历（由近及远，仅限IT相关经历）”**  
       - **“项目经历（与上述工作经历匹配，仅IT相关经历）”**  
    3. 输入你的 **智谱 AI API Key**（可在 [智谱开放平台](https://open.bigmodel.cn/) 获取）  
    4. 可选填写 **供应商缩写、类型、岗位、级别**（若AI未提取到则以此处为准）  
    5. 点击 **开始处理**，等待几秒后下载生成的简历文件  
    """)

col1, col2 = st.columns(2)
with col1:
    pdf_file = st.file_uploader("📂 上传 PDF 简历", type=["pdf"])
with col2:
    excel_template = st.file_uploader("📁 上传 Excel 模板", type=["xlsx"])

with st.container():
    st.subheader("补充信息（若AI未自动识别，请手动填写）")
    col_a, col_b, col_c, col_d = st.columns(4)
    with col_a:
        supplier = st.text_input("供应商缩写 (前4字)", placeholder="例: 腾讯科技 -> 腾讯")
    with col_b:
        emp_type = st.selectbox("类型", ["研发", "测试"])
    with col_c:
        position = st.selectbox("岗位", ["java开发", "前端", "功能测试", "性能测试"])
    with col_d:
        level = st.text_input("级别", placeholder="初级/中级/高级")

if st.button("🚀 开始处理", type="primary"):
    if not pdf_file or not excel_template:
        st.error("请上传 PDF 简历和 Excel 模板")
        st.stop()
    if not api_key:
        st.error("请提供智谱 AI API Key")
        st.stop()
    
    with st.spinner("1/3 正在读取 PDF 内容..."):
        pdf_text = extract_text_from_pdf(pdf_file)
        if not pdf_text.strip():
            st.error("PDF 文本为空，请检查文件是否可解析")
            st.stop()
    
    with st.spinner("2/3 正在调用 AI 分析简历（可能需要10-30秒）..."):
        try:
            ai_result = extract_resume_info(api_key, pdf_text, model=model_name)
            st.success("AI 分析完成")
            with st.expander("查看 AI 提取结果"):
                st.json(ai_result)
        except Exception as e:
            st.error(f"AI 调用失败: {str(e)}")
            st.stop()
    
    # 使用用户补充信息覆盖 extra 字段
    extra = ai_result.get("extra", {})
    if supplier:
        extra["供应商缩写"] = supplier
    if emp_type:
        extra["类型"] = emp_type
    if position:
        extra["岗位"] = position
    if level:
        extra["级别"] = level
    ai_result["extra"] = extra
    
    # 构建输出文件名
    sup = extra.get("供应商缩写", "未知供应商")[:4]
    name = ai_result.get("basic", {}).get("姓名", "未知姓名")
    typ = extra.get("类型", "研发")
    pos = extra.get("岗位", "java开发")
    lvl = extra.get("级别", "")
    filename = f"{sup}-{name}-{typ}-{pos}-{lvl}-简历.xlsx"
    filename = "".join(c for c in filename if c not in r'\/:*?"<>|')
    
    with st.spinner("3/3 正在填充 Excel 模板，保留原格式..."):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_template:
            tmp_template.write(excel_template.getbuffer())
            tmp_template_path = tmp_template.name
        output_path = os.path.join(tempfile.gettempdir(), filename)
        try:
            fill_template(tmp_template_path, output_path, ai_result)
            st.success("填充完成！")
        except Exception as e:
            st.error(f"Excel 填充失败: {str(e)}")
            st.stop()
        finally:
            os.unlink(tmp_template_path)
    
    with open(output_path, "rb") as f:
        st.download_button(
            label="📥 下载生成的简历",
            data=f,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    os.unlink(output_path)
