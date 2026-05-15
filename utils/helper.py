from datetime import datetime

def normalize_date(date_str: str) -> str:
    """将各种日期格式转为 YYYY-MM-DD，缺失日则补当月首日"""
    if not date_str:
        return ""
    # 尝试解析常见格式
    date_str = str(date_str).strip()
    for fmt in ["%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y年%m月%d日", "%Y-%m", "%Y/%m", "%Y.%m", "%Y年%m月"]:
        try:
            if fmt.endswith("%m"):
                # 只有年月
                dt = datetime.strptime(date_str, fmt)
                return dt.strftime("%Y-%m-01")
            else:
                dt = datetime.strptime(date_str, fmt)
                return dt.strftime("%Y-%m-%d")
        except ValueError:
            continue
    # 尝试只提取年份
    import re
    year_match = re.search(r"\b(19|20)\d{2}\b", date_str)
    if year_match:
        return f"{year_match.group()}-01-01"
    return date_str  # 无法解析则原样返回
