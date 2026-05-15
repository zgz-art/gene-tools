import streamlit as st
import os
import tempfile
from utils.pdf_parser import extract_text_from_pdf
from utils.ai_extractor import extract_resume_info
from utils.excel_filler import fill_template

st.set_page_config(page_title="简历智能填充工具", page_icon="📄", layout="wide")

# 自定义CSS提升UI
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
            # 传入模型参数
            ai_result = extract_resume_info(api_key, pdf_text, model=model_name)
            st.success("AI 分析完成")
            # 显示提取结果预览
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
    # 清理非法字符
    filename = "".join(c for c in filename if c not in r'\/:*?"<>|')
    
    with st.spinner("3/3 正在填充 Excel 模板，保留原格式..."):
        # 保存模板到临时文件
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
