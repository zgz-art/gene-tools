import pypdf

def extract_text_from_pdf(pdf_file) -> str:
    """从上传的 PDF 文件中提取纯文本"""
    reader = pypdf.PdfReader(pdf_file)
    text = ""
    for page in reader.pages:
        text += page.extract_text() + "\n"
    return text
