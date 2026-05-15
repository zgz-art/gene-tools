from zhipuai import ZhipuAI
import json

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
