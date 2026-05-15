import streamlit as st
import os
import tempfile
import json
from datetime import datetime
import re
import pypdf
from zhipuai import ZhipuAI
import openpyxl
from copy import copy

# ==================== 页面配置 ====================
st.set_page_config(page_title="简历智能填充工具", page_icon="📄", layout="wide")

st.markdown("""
<style>
    .stButton button { background-color: #4CAF50; color: white; border-radius: 8px; }
    h1 { color: #2c3e50; text-align: center; }
</style>
""", unsafe_allow_html=True)

# ==================== 初始化 Session State ====================
if "pdf_content" not in st.session_state:
    st.session_state.pdf_content = None
if "excel_content" not in st.session_state:
    st.session_state.excel_content = None
if "ai_result" not in st.session_state:
    st.session_state.ai_result = None
if "pdf_text" not in st.session_state:
    st.session_state.pdf_text = None
if "last_pdf_name" not in st.session_state:
    st.session_state.last_pdf_name = None
if "last_excel_name" not in st.session_state:
    st.session_state.last_excel_name = None

# ==================== 辅助函数 ====================
def normalize_date(date_str: str) -> str:
    if not date_str:
        return ""
    date_str = str(date_str).strip()
    for fmt in ["%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y年%m月%d日", "%Y-%m", "%Y/%m", "%Y.%m", "%Y年%m月"]:
        try:
            if fmt.endswith("%m"):
                return datetime.strptime(date_str, fmt).strftime("%Y-%m-01")
            else:
                return datetime.strptime(date_str, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    year_match = re.search(r"\b(19|20)\d{2}\b", date_str)
    return f"{year_match.group()}-01-01" if year_match else date_str

def extract_text_from_pdf(pdf_file) -> str:
    reader = pypdf.PdfReader(pdf_file)
    return "\n".join(page.extract_text() for page in reader.pages)

def get_primary_cell(worksheet, row, col):
    """获取合并单元格的主单元格（左上角）"""
    for merged in worksheet.merged_cells.ranges:
        if (merged.min_row <= row <= merged.max_row and
            merged.min_col <= col <= merged.max_col):
            return worksheet.cell(merged.min_row, merged.min_col)
    return worksheet.cell(row, col)

def safe_write(worksheet, row, col, value):
    """安全写入值，自动处理合并单元格"""
    get_primary_cell(worksheet, row, col).value = value

def normalize_string(s: str) -> str:
    """标准化字符串：去除空格、冒号、句号等标点，用于模糊匹配"""
    if not s:
        return ""
    s = str(s).strip()
    s = re.sub(r'[：:，,。、；;]', '', s)
    s = re.sub(r'\s+', '', s)
    return s

def find_cell(worksheet, value, exact=False):
    norm_value = normalize_string(value)
    for row in worksheet.iter_rows():
        for cell in row:
            if cell.value is None:
                continue
            if exact:
                if normalize_string(cell.value) == norm_value:
                    return cell.row, cell.column
            else:
                if norm_value in normalize_string(cell.value):
                    return cell.row, cell.column
    return None

def find_row_by_keyword(worksheet, keyword):
    norm_keyword = normalize_string(keyword)
    for row in range(1, worksheet.max_row + 1):
        for col in range(1, worksheet.max_column + 1):
            val = worksheet.cell(row, col).value
            if val and norm_keyword in normalize_string(val):
                return row
    return None

def clear_row_content(worksheet, row, max_col):
    for col in range(1, max_col + 1):
        primary = get_primary_cell(worksheet, row, col)
        if primary.row == row and primary.column == col:
            primary.value = None

# ==================== 填充逻辑 ====================
def fill_basic_info(ws, basic_data):
    for key, value in basic_data.items():
        if not value:
            continue
        pos = find_cell(ws, key, exact=False)
        if pos:
            safe_write(ws, pos[0], pos[1] + 1, value)

def fill_education_block(ws, keyword, edu_data, undergrad_uni_cell=None, undergrad_major_cell=None, grad_uni_cell=None, grad_major_cell=None):
    """
    填充教育信息。
    对于本科和研究生，毕业院校和专业使用绝对坐标写入；其他字段模糊匹配。
    """
    pos = find_cell(ws, keyword, exact=True)
    if not pos:
        st.warning(f"未找到关键字: {keyword}")
        return
    row_start = pos[0] + 1
    # 确定当前是本科还是研究生
    is_undergrad = "本科" in keyword
    for offset in range(9):
        current_row = row_start + offset
        field_cell = ws.cell(current_row, 1)
        if field_cell.value is None:
            if offset > 10:
                break
            continue
        field_name = normalize_string(field_cell.value)
        # 特殊处理毕业院校和专业：使用绝对坐标直接写入
        if is_undergrad and undergrad_uni_cell:
            if normalize_string("毕业院校") == field_name or normalize_string("毕业学校") == field_name:
                uni_val = edu_data.get("毕业院校", "")
                if uni_val:
                    safe_write(ws, undergrad_uni_cell[0], undergrad_uni_cell[1], uni_val)
                continue  # 跳过后续模糊匹配
            if normalize_string("专业") == field_name and undergrad_major_cell:
                major_val = edu_data.get("专业", "")
                if major_val:
                    safe_write(ws, undergrad_major_cell[0], undergrad_major_cell[1], major_val)
                continue
        elif not is_undergrad and grad_uni_cell:
            if normalize_string("毕业院校") == field_name or normalize_string("毕业学校") == field_name:
                uni_val = edu_data.get("毕业院校", "")
                if uni_val:
                    safe_write(ws, grad_uni_cell[0], grad_uni_cell[1], uni_val)
                continue
            if normalize_string("专业") == field_name and grad_major_cell:
                major_val = edu_data.get("专业", "")
                if major_val:
                    safe_write(ws, grad_major_cell[0], grad_major_cell[1], major_val)
                continue
        
        # 其他字段模糊匹配（如入学时间、毕业时间等）
        for data_key, data_val in edu_data.items():
            if data_val and normalize_string(data_key) == field_name:
                safe_write(ws, current_row, 2, data_val)
                break

def fill_work_experience(ws, work_list):
    keyword_row = find_row_by_keyword(ws, "工作经历（由近及远，仅限IT相关经历）")
    if not keyword_row:
        st.warning("未找到工作经历关键字，跳过填充")
        return
    header_row = keyword_row + 1
    data_start = header_row + 1
    reserved = 3  # 固定预留3行，不动态增行
    max_col = ws.max_column

    # 清空预留区域
    for r in range(data_start, data_start + reserved):
        clear_row_content(ws, r, max_col)

    sorted_work = sorted(work_list, key=lambda x: x.get("开始日期", "1900-01-01"), reverse=True)
    # 只填充前 reserved 条，超出部分提示警告
    if len(sorted_work) > reserved:
        st.warning(f"工作经历共有 {len(sorted_work)} 条，模板仅预留 {reserved} 行，超出部分将被忽略。")
        sorted_work = sorted_work[:reserved]

    for idx, work in enumerate(sorted_work):
        target_row = data_start + idx
        for col in range(1, max_col + 1):
            header = ws.cell(header_row, col).value
            if not header:
                continue
            header_norm = normalize_string(header)
            if "开始日期" in header_norm:
                safe_write(ws, target_row, col, normalize_date(work.get("开始日期", "")))
            elif "结束日期" in header_norm:
                safe_write(ws, target_row, col, normalize_date(work.get("结束日期", "")))
            elif "单位名称" in header_norm:
                safe_write(ws, target_row, col, work.get("单位名称", ""))
            elif "岗位" in header_norm:
                safe_write(ws, target_row, col, work.get("岗位", ""))

def fill_project_experience(ws, project_list):
    keyword_row = find_row_by_keyword(ws, "项目经历（与上述工作经历匹配，仅IT相关经历）")
    if not keyword_row:
        st.warning("未找到项目经历关键字，跳过填充")
        return
    header_row = keyword_row + 1
    data_start = header_row + 1
    reserved = 6  # 固定预留6行
    max_col = ws.max_column

    for r in range(data_start, data_start + reserved):
        clear_row_content(ws, r, max_col)

    sorted_proj = sorted(project_list, key=lambda x: x.get("开始日期", "1900-01-01"), reverse=True)
    if len(sorted_proj) > reserved:
        st.warning(f"项目经历共有 {len(sorted_proj)} 条，模板仅预留 {reserved} 行，超出部分将被忽略。")
        sorted_proj = sorted_proj[:reserved]

    for idx, proj in enumerate(sorted_proj):
        target_row = data_start + idx
        for col in range(1, max_col + 1):
            header = ws.cell(header_row, col).value
            if not header:
                continue
            header_norm = normalize_string(header)
            if "开始日期" in header_norm:
                safe_write(ws, target_row, col, normalize_date(proj.get("开始日期", "")))
            elif "结束日期" in header_norm:
                safe_write(ws, target_row, col, normalize_date(proj.get("结束日期", "")))
            elif "项目名称" in header_norm:
                safe_write(ws, target_row, col, proj.get("项目名称", ""))
            elif "项目描述" in header_norm:
                safe_write(ws, target_row, col, proj.get("项目描述", ""))
            elif "项目角色" in header_norm:
                safe_write(ws, target_row, col, proj.get("项目角色", ""))

def fill_template(template_path, output_path, ai_data):
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active
    
    # 基础信息
    fill_basic_info(ws, ai_data.get("basic", {}))
    
    # 本科学历：指定毕业院校 D14，专业 D15
    undergrad_data = ai_data.get("education", {}).get("undergraduate", {})
    if undergrad_data:
        fill_education_block(ws, "本科学历", undergrad_data, 
                             undergrad_uni_cell=(14, 4),   # D14 = 第14行，第4列
                             undergrad_major_cell=(15, 4)) # D15 = 第15行，第4列
    
    # 研究生学历：指定毕业院校 D23，专业 D24
    postgrad_data = ai_data.get("education", {}).get("postgraduate", {})
    if postgrad_data:
        fill_education_block(ws, "研究生学历", postgrad_data,
                             grad_uni_cell=(23, 4),   # D23
                             grad_major_cell=(24, 4)) # D24
    
    # 工作经历和项目经历
    if ai_data.get("work_experience"):
        fill_work_experience(ws, ai_data["work_experience"])
    if ai_data.get("project_experience"):
        fill_project_experience(ws, ai_data["project_experience"])
    
    wb.save(output_path)

# ==================== AI 提取 ====================
SYSTEM_PROMPT = """你是一个专业的简历信息提取助手。请从以下简历文本中提取指定字段，并以 JSON 格式返回。
要求：
1. 日期统一使用 "YYYY-MM-DD" 格式，若只提供年月，则默认补充 "-01"。
2. 若某字段不存在，返回空字符串或空列表。
3. 工作经历和项目经历请按时间由近及远排序（开始日期越晚越靠前）。

输出 JSON 结构如下：
{
    "basic": {
        "姓名": "", "身份证号": "", "出生日期": "", "电话": "",
        "首次参加工作时间": "", "首次参加IT领域工作时间": "", "最高学历": "",
        "掌握语言": "", "掌握技能": "", "专业证书": ""
    },
    "education": {
        "undergraduate": {
            "入学时间": "", "毕业院校": "", "毕业时间": "", "专业": "",
            "毕业证编号": "", "毕业证学信网在线验证码": "",
            "学位证编号": "", "学位证学信网在线验证码": ""
        },
        "postgraduate": {
            "入学时间": "", "毕业院校": "", "毕业时间": "", "专业": "",
            "毕业证编号": "", "毕业证学信网在线验证码": "",
            "学位证编号": "", "学位证学信网在线验证码": ""
        }
    },
    "work_experience": [
        {"开始日期": "", "结束日期": "", "单位名称": "", "岗位": ""}
    ],
    "project_experience": [
        {"开始日期": "", "结束日期": "", "项目名称": "", "项目描述": "", "项目角色": ""}
    ],
    "extra": {"供应商缩写": "", "类型": "", "岗位": "", "级别": ""}
}
"""

def extract_resume_info(api_key: str, pdf_text: str, model: str = "glm-4-plus") -> dict:
    client = ZhipuAI(api_key=api_key)
    resp = client.chat.completions.create(
        model=model,
        messages=[{"role": "system", "content": SYSTEM_PROMPT},
                  {"role": "user", "content": f"简历文本：\n{pdf_text}"}],
        temperature=0.1,
        response_format={"type": "json_object"}
    )
    return json.loads(resp.choices[0].message.content)

# ==================== Streamlit UI ====================
st.title("📄 智能简历填充工具")
st.markdown("上传 PDF 简历和 Excel 模板，AI 自动提取信息并生成标准格式文件。")

with st.sidebar:
    st.header("⚙️ 配置")
    api_key = st.text_input("🔑 智谱 AI API Key", type="password", help="必填，用于调用智谱AI")
    model_name = st.selectbox("🤖 选择大模型", ["glm-4-plus", "glm-4-flash", "glm-4-air", "glm-4-long"], index=0)
    st.markdown("---")
    st.caption("模板中请包含：本科学历、研究生学历、工作经历（...）、项目经历（...）等关键字")

with st.expander("📌 使用说明", expanded=True):
    st.markdown("""
    1. 准备 **PDF 格式** 的原始简历文件  
    2. 准备 **Excel 模板**（.xlsx），模板中应包含以下关键字：  
       - 基础信息字段（姓名、身份证号...）  
       - 精确的 **"本科学历"**、**"研究生学历"** 文字  
       - **"工作经历（由近及远，仅限IT相关经历）"**  
       - **"项目经历（与上述工作经历匹配，仅IT相关经历）"**  
    3. 输入你的 **智谱 AI API Key**（可在 [智谱开放平台](https://open.bigmodel.cn/) 获取）  
    4. 可选填写 **供应商缩写、类型、岗位、级别**（若AI未提取到则以此处为准）  
    5. 点击 **开始处理**，等待几秒后下载生成的简历文件  
    **注意**：工作经历模板预留3行，项目经历预留6行，超出部分会被忽略。
    """)

col1, col2 = st.columns(2)
with col1:
    pdf_file = st.file_uploader("📂 PDF 简历", type=["pdf"])
with col2:
    excel_template = st.file_uploader("📁 Excel 模板", type=["xlsx"])

# 检测文件是否变化
if pdf_file is not None and st.session_state.last_pdf_name != pdf_file.name:
    st.session_state.ai_result = None
    st.session_state.pdf_text = None
    st.session_state.last_pdf_name = pdf_file.name
if excel_template is not None and st.session_state.last_excel_name != excel_template.name:
    st.session_state.ai_result = None
    st.session_state.last_excel_name = excel_template.name

with st.container():
    st.subheader("补充信息（优先使用此处填写）")
    col_a, col_b, col_c, col_d = st.columns(4)
    with col_a:
        supplier = st.text_input("供应商缩写 (前4字)", placeholder="例如：腾讯")
    with col_b:
        emp_type = st.selectbox("类型", ["研发", "测试"])
    with col_c:
        position = st.selectbox("岗位", ["java开发", "前端", "功能测试", "性能测试"])
    with col_d:
        level = st.text_input("级别", placeholder="初级/中级/高级")

# 校验必填项
if st.button("🚀 开始处理", type="primary"):
    errors = []
    if not pdf_file:
        errors.append("请上传 PDF 简历文件")
    if not excel_template:
        errors.append("请上传 Excel 模板文件")
    if not api_key:
        errors.append("请输入智谱 AI API Key")
    
    if errors:
        for err in errors:
            st.error(f"❌ {err}")
        st.stop()
    
    # 读取 PDF 内容
    if st.session_state.pdf_text is None:
        with st.spinner("读取 PDF..."):
            st.session_state.pdf_text = extract_text_from_pdf(pdf_file)
            if not st.session_state.pdf_text.strip():
                st.error("❌ PDF 文本为空，请检查文件是否可解析")
                st.stop()

    with st.spinner(f"调用 {model_name} 分析简历（可能需要30秒）..."):
        try:
            ai_result = extract_resume_info(api_key, st.session_state.pdf_text, model_name)
            st.session_state.ai_result = ai_result
            st.success("✅ AI 分析完成")
        except Exception as e:
            st.error(f"❌ AI 调用失败: {e}")
            st.stop()

# 显示 AI 提取结果
if st.session_state.ai_result is not None:
    with st.expander("查看 AI 提取结果", expanded=True):
        st.json(st.session_state.ai_result)

# 下载按钮
if st.session_state.ai_result is not None and excel_template is not None:
    # 合并补充信息
    extra = st.session_state.ai_result.get("extra", {})
    if supplier:
        extra["供应商缩写"] = supplier
    if emp_type:
        extra["类型"] = emp_type
    if position:
        extra["岗位"] = position
    if level:
        extra["级别"] = level
    st.session_state.ai_result["extra"] = extra

    sup = extra.get("供应商缩写", "未知")[:4]
    name = st.session_state.ai_result.get("basic", {}).get("姓名", "未知")
    typ = extra.get("类型", "研发")
    pos = extra.get("岗位", "java开发")
    lvl = extra.get("级别", "")
    filename = f"{sup}-{name}-{typ}-{pos}-{lvl}-简历.xlsx"
    filename = "".join(c for c in filename if c not in r'\/:*?"<>|')

    if st.button("📥 下载生成的简历", type="secondary"):
        with st.spinner("填充 Excel（完全保留原格式）..."):
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                tmp.write(excel_template.getbuffer())
                tmp_path = tmp.name
            out_path = os.path.join(tempfile.gettempdir(), filename)
            try:
                fill_template(tmp_path, out_path, st.session_state.ai_result)
                st.success("✅ 填充成功！")
            except Exception as e:
                st.error(f"❌ Excel 填充失败: {e}")
                st.stop()
            finally:
                os.unlink(tmp_path)

        with open(out_path, "rb") as f:
            st.download_button("📥 点击下载文件", f, file_name=filename, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        os.unlink(out_path)
