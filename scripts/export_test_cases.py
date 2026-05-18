#!/usr/bin/env python
import argparse
import json
import re
from collections import Counter
from copy import deepcopy
from datetime import datetime
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# 边界值策略关键词
BOUNDARY_KEYWORDS = {
    "金额": ["元", "¥", "$", "USD", "CNY", "金额", "价格", "费用", "成本", "预算"],
    "百分比": ["%", "百分比", "比例", "折扣", "费率", "利率"],
    "时间": ["秒", "分", "小时", "天", "毫秒", "超时", "有效期", "过期"],
    "文件": ["B", "KB", "MB", "GB", "文件大小", "上传限制"],
    "字符": ["位", "字符", "字", "长度", "字数", "字节"],
    "整数": ["条", "个", "次", "笔", "件", "条数", "数量", "次数", "页数"],
}

# 优先级顺序
KEYWORD_PRIORITY = ["金额", "百分比", "时间", "文件", "字符", "整数"]


def _parse_number(num_str):
    """智能解析数字，整数保持整数类型，小数保持小数类型"""
    try:
        num = float(num_str)
        if num.is_integer():
            return int(num)
        return num
    except:
        return num_str


def extract_boundary_values(text):
    """从文本中提取边界值信息"""
    boundary_info = []
    
    # 匹配数字范围，如 "10-100"、"0-9999"、"18-60周岁"
    range_pattern = r'(\d+(?:\.\d+)?)\s*[-~至到]\s*(\d+(?:\.\d+)?)'
    ranges = re.findall(range_pattern, text)
    for match in ranges:
        min_val = _parse_number(match[0])
        max_val = _parse_number(match[1])
        unit = _detect_unit(text, match[0], match[1])
        boundary_info.append({
            "type": "range",
            "min": min_val,
            "max": max_val,
            "unit": unit,
            "text": f"{match[0]}-{match[1]}"
        })
    
    # 匹配单个数值限制，如 "最多100条"、"上限5000元"、"不超过10MB"
    limit_pattern = r'(最多|最少|上限|下限|不超过|不少于|不低于|不高于|超过|小于|大于)\s*(\d+(?:\.\d+)?)([a-zA-Z\u4e00-\u9fa5]*)'
    limits = re.findall(limit_pattern, text)
    for match in limits:
        operator = match[0]
        value = _parse_number(match[1])
        unit = match[2] if match[2] else _detect_unit(text, match[1])
        boundary_info.append({
            "type": "limit",
            "operator": operator,
            "value": value,
            "unit": unit,
            "text": f"{operator}{match[1]}{unit}"
        })
    
    return boundary_info


def _detect_unit(text, num_str, end_str=None):
    """检测数值的单位"""
    for category, keywords in BOUNDARY_KEYWORDS.items():
        for keyword in keywords:
            if keyword in text:
                return keyword
    return ""


def generate_boundary_test_cases(module_name, prd_content, existing_cases):
    """根据PRD内容生成边界值测试用例"""
    boundary_info = extract_boundary_values(prd_content)
    new_cases = []
    existing_titles = {case.get("用例标题", "") for case in existing_cases}
    
    # 获取最大用例编号
    max_num = 0
    for case in existing_cases:
        match = re.search(r'TC-\w+-\d+', case.get("用例编号", ""))
        if match:
            try:
                num = int(match.group().split('-')[-1])
                max_num = max(max_num, num)
            except:
                pass
    
    for info in boundary_info:
        if info["type"] == "range":
            title = f"边界值测试：{info['text']}{info['unit']}范围"
            if title not in existing_titles:
                max_num += 1
                case = {
                    "用例编号": f"TC-{module_name[:3].upper()}-{max_num:03d}",
                    "模块": module_name,
                    "用例标题": title,
                    "前置条件": "系统正常运行，用户已登录",
                    "测试步骤": f"""1. 设置值为{info['min'] - 1}{info['unit']}进行操作
2. 设置值为{info['min']}{info['unit']}进行操作
3. 设置值为{info['min'] + 1}{info['unit']}进行操作
4. 设置值为{info['max'] - 1}{info['unit']}进行操作
5. 设置值为{info['max']}{info['unit']}进行操作
6. 设置值为{info['max'] + 1}{info['unit']}进行操作""",
                    "预期结果": f"""1. 小于{info['min']}{info['unit']}时应提示超出范围
2. 等于{info['min']}{info['unit']}时应正常处理
3. 大于{info['min']}{info['unit']}时应正常处理
4. 小于{info['max']}{info['unit']}时应正常处理
5. 等于{info['max']}{info['unit']}时应正常处理
6. 大于{info['max']}{info['unit']}时应提示超出范围""",
                    "实际结果": "",
                    "优先级": "中",
                    "类型": "边界值测试",
                    "备注": f"边界值范围: {info['min']}-{info['max']}{info['unit']}"
                }
                new_cases.append(case)
        
        elif info["type"] == "limit":
            title = f"边界值测试：{info['text']}"
            if title not in existing_titles:
                max_num += 1
                case = {
                    "用例编号": f"TC-{module_name[:3].upper()}-{max_num:03d}",
                    "模块": module_name,
                    "用例标题": title,
                    "前置条件": "系统正常运行，用户已登录",
                    "测试步骤": f"""1. 设置值为{info['value'] - 1}{info['unit']}进行操作
2. 设置值为{info['value']}{info['unit']}进行操作
3. 设置值为{info['value'] + 1}{info['unit']}进行操作""",
                    "预期结果": f"""1. 小于{info['value']}{info['unit']}时应正常处理
2. 等于{info['value']}{info['unit']}时应正常处理
3. 大于{info['value']}{info['unit']}时应提示超出限制""",
                    "实际结果": "",
                    "优先级": "中",
                    "类型": "边界值测试",
                    "备注": f"限制值: {info['value']}{info['unit']}"
                }
                new_cases.append(case)
    
    return new_cases

DEFAULT_COLUMNS = [
    "用例编号",
    "模块",
    "用例标题",
    "前置条件",
    "测试步骤",
    "预期结果",
    "实际结果",
    "优先级",
    "类型",
    "备注",
]

DEFAULT_OUTPUT_DIR = Path.cwd()

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(color="FFFFFF", bold=True)
WRAP_ALIGNMENT = Alignment(wrap_text=True, vertical="top")
INVALID_SHEET_CHARS = '[]:*?/\\'

def parse_args():
    parser = argparse.ArgumentParser(
        description="Export categorized test cases to an xlsx workbook."
    )
    parser.add_argument("--input-json", required=True, help="Path to structured test-case JSON.")
    parser.add_argument("--output-dir", help="Output directory for the workbook.")
    parser.add_argument("--file-name", help="Optional workbook file name.")
    return parser.parse_args()

def sanitize_filename_part(value):
    if not value:
        return ""
    invalid = '<>:"/\\|?*'
    sanitized = "".join("_" if ch in invalid else ch for ch in str(value).strip())
    return sanitized.strip(" ._")

def build_default_name(payload):
    today = datetime.now().strftime("%Y-%m-%d")
    system_name = sanitize_filename_part(payload.get("system_name", ""))
    module_name = sanitize_filename_part(payload.get("module_name", ""))
    if system_name and module_name:
        return f"{system_name}_{module_name}_测试用例_{today}.xlsx"
    if module_name:
        return f"{module_name}_测试用例_{today}.xlsx"
    return f"测试用例_{today}.xlsx"

def unique_path(directory, file_name):
    base = Path(directory)
    candidate = base / file_name
    if not candidate.exists():
        return candidate
    stem = candidate.stem
    suffix = candidate.suffix
    index = 1
    while True:
        next_candidate = base / f"{stem}_{index:02d}{suffix}"
        if not next_candidate.exists():
            return next_candidate
        index += 1

def normalize_payload(payload):
    normalized = deepcopy(payload)
    normalized.setdefault("system_name", "")
    normalized.setdefault("module_name", "")
    normalized.setdefault("input_summary", "")
    normalized.setdefault("generated_at", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    normalized.setdefault("output_format", "xlsx")
    normalized.setdefault("uncertainties", [])
    normalized.setdefault("extra_columns", [])
    normalized.setdefault("categories", [])
    return normalized

def build_workbook(payload):
    payload = normalize_payload(payload)
    workbook = Workbook()
    summary = workbook.active
    summary.title = "说明"
    extra_columns = payload.get("extra_columns", [])
    all_columns = DEFAULT_COLUMNS + [c for c in extra_columns if c not in DEFAULT_COLUMNS]
    category_counts = Counter()
    total_cases = 0
    
    # 收集所有现有用例（用于去重）
    all_existing_cases = []
    for category in payload["categories"]:
        all_existing_cases.extend(category.get("cases", []))
    
    # 从PRD内容中自动生成边界值测试用例
    prd_content = payload.get("input_summary", "")
    module_name = payload.get("module_name", "模块")
    auto_boundary_cases = generate_boundary_test_cases(module_name, prd_content, all_existing_cases)
    
    # 查找或创建边界值分类
    boundary_category = None
    for category in payload["categories"]:
        if category.get("name", "").strip() == "边界值":
            boundary_category = category
            break
    
    if boundary_category:
        # 将自动生成的边界值用例添加到现有边界值分类
        boundary_category["cases"].extend(auto_boundary_cases)
    elif auto_boundary_cases:
        # 如果没有边界值分类，则创建新分类
        payload["categories"].append({
            "name": "边界值",
            "cases": auto_boundary_cases
        })
    
    for category in payload["categories"]:
        category_name = str(category.get("name", "")).strip() or "未分类"
        cases = category.get("cases", [])
        category_counts[category_name] += len(cases)
        total_cases += len(cases)
        sheet = workbook.create_sheet(title=trim_sheet_name(category_name))
        write_case_sheet(sheet, all_columns, cases)
    write_summary_sheet(summary, payload, category_counts, total_cases)
    return workbook

def trim_sheet_name(name):
    cleaned = "".join("_" if ch in INVALID_SHEET_CHARS else ch for ch in name)
    return cleaned[:31] or "Sheet"

def write_case_sheet(sheet, columns, cases):
    sheet.append(columns)
    style_header_row(sheet, 1)
    sheet.freeze_panes = "A2"
    for case in cases:
        row = [case.get(column, "") for column in columns]
        sheet.append(row)
    apply_case_sheet_formatting(sheet)

def write_summary_sheet(sheet, payload, category_counts, total_cases):
    rows = [
        ("系统名", payload.get("system_name", "")),
        ("模块名", payload.get("module_name", "")),
        ("输入来源", payload.get("input_summary", "")),
        ("生成时间", payload.get("generated_at", "")),
        ("输出格式", payload.get("output_format", "xlsx")),
        ("用例总数", total_cases),
    ]
    for label, value in rows:
        sheet.append([label, value])
    sheet.append([])
    sheet.append(["分类", "数量"])
    style_header_row(sheet, sheet.max_row)
    for name, count in category_counts.items():
        sheet.append([name, count])
    sheet.append([])
    sheet.append(["待确认项"])
    style_header_row(sheet, sheet.max_row)
    uncertainties = payload.get("uncertainties", [])
    if uncertainties:
        for item in uncertainties:
            sheet.append([item])
    else:
        sheet.append(["无"])
    apply_summary_formatting(sheet)

def style_header_row(sheet, row_index):
    for cell in sheet[row_index]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = WRAP_ALIGNMENT

def apply_case_sheet_formatting(sheet):
    for row in sheet.iter_rows():
        for cell in row:
            cell.alignment = WRAP_ALIGNMENT
    max_width = 40
    min_width = 12
    for column_cells in sheet.columns:
        letter = get_column_letter(column_cells[0].column)
        max_len = max(len(str(cell.value or "")) for cell in column_cells)
        sheet.column_dimensions[letter].width = max(min_width, min(max_width, max_len + 2))

def apply_summary_formatting(sheet):
    for row in sheet.iter_rows():
        for cell in row:
            cell.alignment = WRAP_ALIGNMENT
    sheet.column_dimensions["A"].width = 18
    sheet.column_dimensions["B"].width = 60

def main():
    args = parse_args()
    input_path = Path(args.input_json)
    with input_path.open("r", encoding="utf-8") as fh:
        payload = json.load(fh)
    output_dir = Path(args.output_dir or payload.get("output_dir") or DEFAULT_OUTPUT_DIR)
    output_dir.mkdir(parents=True, exist_ok=True)
    file_name = args.file_name or payload.get("file_name") or build_default_name(payload)
    if not file_name.lower().endswith(".xlsx"):
        file_name = f"{file_name}.xlsx"
    workbook = build_workbook(payload)
    final_path = unique_path(output_dir, sanitize_filename_part(file_name[:-5]) + ".xlsx")
    workbook.save(final_path)
    print(str(final_path))

if __name__ == "__main__":
    main()
