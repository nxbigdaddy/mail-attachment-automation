import os
import re
from datetime import datetime
from docx import Document
from openpyxl import load_workbook

def read_docx_full(filepath):
    try:
        doc = Document(filepath)
        return "\n".join([p.text for p in doc.paragraphs])
    except:
        return ""

def read_xlsx_full(filepath):
    try:
        wb = load_workbook(filepath)
        sheet = wb.active
        lines = []
        for row in sheet.iter_rows(values_only=True):
            row_content = [str(cell) if cell else "" for cell in row]
            if any(cell.strip() for cell in row_content):
                lines.append(" | ".join(row_content))
        return "\n".join(lines)
    except:
        return ""

def extract_meaningful_work(content):
    lines = content.split('\n')
    work_items = []
    skip_keywords = ['工 作 日 报', '汇 报 人', '序号', '项目名称', '栏目', '预计工时', '实际工时',
                     '预计进度', '实际进度', '作业备注', '明 日 作 业 计 划', '作 业 问 题 报 告',
                     '今 日 作 业 情 况', '日期', '星期', '年 月 日']

    for line in lines:
        line = line.strip()
        if not line or len(line) < 10:
            continue

        should_skip = any(kw in line for kw in skip_keywords)
        if should_skip:
            continue

        cleaned = re.sub(r'^[\d\.\、\-\)]+\s*', '', line)
        cleaned = re.sub(r'[\|\s]{2,}', ' ', cleaned)
        cleaned = cleaned.strip()

        if len(cleaned) > 15 and not cleaned.startswith('1 |'):
            work_items.append(cleaned)

    seen = set()
    unique_items = []
    for item in work_items:
        if item not in seen and len(item) > 15:
            seen.add(item)
            unique_items.append(item)

    return unique_items[:6]

def parse_date(filename):
    patterns = [
        r'(\d{4})[-/年](\d{1,2})[-/月](\d{1,2})',
        r'-(\d{4})-(\d{2})-(\d{2})',
        r'(\d{4})-(\d{2})-(\d{2})',
        r'(\d{2})\.(\d{1,2})\.(\d{1,2})',
        r'-(\d{2})\.(\d{1,2})\.(\d{1,2})',
    ]

    for pattern in patterns:
        match = re.search(pattern, filename)
        if match:
            groups = match.groups()
            try:
                if len(groups[0]) == 4:
                    year, month, day = int(groups[0]), int(groups[1]), int(groups[2])
                else:
                    year, month, day = 2000 + int(groups[0]), int(groups[1]), int(groups[2])

                if 1 <= month <= 12 and 1 <= day <= 31:
                    return datetime(year, month, day)
            except:
                pass
    return None

def generate_clean_summary_html(dinghai_data, shuibin_data):
    html = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>下属工作汇总报告</title>
    <style>
        @page {
            size: A4 landscape;
            margin: 1.5cm;
        }
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        body {
            font-family: 'Microsoft YaHei', 'SimHei', sans-serif;
            font-size: 10pt;
            line-height: 1.8;
            color: #333;
            background: #fff;
        }
        .page {
            page-break-after: always;
            min-height: 100vh;
            padding: 30px;
        }
        .page:last-child {
            page-break-after: avoid;
        }
        .cover {
            text-align: center;
            padding-top: 100px;
        }
        .cover h1 {
            font-size: 30pt;
            color: #1a365d;
            margin-bottom: 25px;
        }
        .cover .subtitle {
            font-size: 15pt;
            color: #4a5568;
            margin-bottom: 50px;
        }
        .cover .meta {
            font-size: 12pt;
            color: #718096;
        }
        .main-content {
            margin-top: 60px;
        }
        .person-section {
            margin-bottom: 50px;
            page-break-inside: avoid;
        }
        .person-header {
            background: linear-gradient(135deg, #f7fafc 0%, #edf2f7 100%);
            padding: 20px 25px;
            border-radius: 10px;
            margin-bottom: 20px;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        .person-header h2 {
            font-size: 18pt;
            color: #2d3748;
        }
        .person-header .days {
            font-size: 11pt;
            color: #718096;
            background: #fff;
            padding: 5px 15px;
            border-radius: 20px;
        }
        .work-category {
            margin-bottom: 20px;
        }
        .work-category h3 {
            font-size: 13pt;
            color: #2d3748;
            margin-bottom: 12px;
            padding-left: 12px;
            border-left: 4px solid #3182ce;
        }
        .work-item {
            background: #f8fafc;
            padding: 12px 18px;
            margin-bottom: 8px;
            border-radius: 6px;
            border-left: 3px solid #4299e1;
            font-size: 10pt;
            color: #4a5568;
        }
        .highlights-section {
            margin-top: 60px;
            page-break-inside: avoid;
        }
        .highlights-section h2 {
            font-size: 18pt;
            color: #1a365d;
            text-align: center;
            margin-bottom: 30px;
        }
        .highlight-grid {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 30px;
        }
        .highlight-card {
            background: linear-gradient(135deg, #ebf8ff 0%, #e2e8f0 100%);
            border-radius: 12px;
            padding: 25px;
        }
        .highlight-card h3 {
            font-size: 15pt;
            color: #2d3748;
            margin-bottom: 15px;
        }
        .highlight-card ul {
            list-style: none;
            padding: 0;
        }
        .highlight-card li {
            padding: 8px 0;
            font-size: 10pt;
            color: #4a5568;
            border-bottom: 1px solid #e2e8f0;
        }
        .highlight-card li:last-child {
            border-bottom: none;
        }
        .highlight-card li strong {
            color: #2d3748;
        }
        .footer {
            margin-top: 60px;
            padding-top: 20px;
            border-top: 2px solid #e2e8f0;
            text-align: center;
            font-size: 9pt;
            color: #a0aec0;
        }
        @media print {
            body {
                font-size: 9pt;
            }
        }
    </style>
</head>
<body>
    <div class="page">
        <div class="cover">
            <h1>📊 下属工作汇总报告</h1>
            <div class="subtitle">汇总周期：2026年5月18日 - 2026年5月26日</div>
            <div class="meta">生成时间：""" + datetime.now().strftime("%Y年%m月%d日") + """</div>
        </div>
    </div>

    <div class="page">
        <div class="main-content">
"""

    for person, data, color in [("示例员工", dinghai_data, "#3182ce"), ("水彬", shuibin_data, "#38a169")]:
        html += f"""            <div class="person-section">
                <div class="person-header" style="border-left: 5px solid {color};">
                    <h2>👤 {person}</h2>
                    <div class="days">📅 {len(data['works'])} 个工作日</div>
                </div>
"""

        if data.get('works'):
            html += """                <div class="work-category">
                    <h3>📋 主要工作内容</h3>
"""
            for work in data['works'][:8]:
                html += f"""                    <div class="work-item">{work}</div>
"""
            html += """                </div>
"""

        html += """            </div>
"""

    html += """        </div>

        <div class="footer">
            本报告由腾讯企业邮箱自动生成 | 仅供参考
        </div>
    </div>

    <div class="page">
        <div class="highlights-section">
            <h2>⭐ 工作重点成果</h2>

            <div class="highlight-grid">
"""

    for person, data, color in [("示例员工", dinghai_data, "#3182ce"), ("水彬", shuibin_data, "#38a169")]:
        html += f"""                <div class="highlight-card">
                    <h3 style="color: {color};">{person}</h3>
                    <ul>
"""
        if data.get('highlights'):
            for h in data['highlights']:
                html += f"                        <li><strong>✓</strong> {h}</li>\n"
        html += """                    </ul>
                </div>
"""

    html += """            </div>
        </div>

        <div class="footer">
            报告生成时间：""" + datetime.now().strftime("%Y年%m月%d日") + """ | 腾讯企业邮箱自动汇总
        </div>
    </div>
</body>
</html>"""

    return html

def main():
    base_dir = r"<workspace>\mail-auto\attachments_20260531_132532"
    start_date = datetime(2026, 5, 18)
    end_date = datetime(2026, 5, 26)

    dinghai_data = {"works": [], "highlights": []}
    shuibin_data = {"works": [], "highlights": []}

    persons = {
        "示例员工": dinghai_data,
        "水彬": shuibin_data
    }

    for person_dir in ["示例员工", "水彬"]:
        full_dir = os.path.join(base_dir, person_dir)
        if not os.path.exists(full_dir):
            continue

        for filename in os.listdir(full_dir):
            doc_date = parse_date(filename)
            if not doc_date or not (start_date <= doc_date <= end_date):
                continue

            filepath = os.path.join(full_dir, filename)
            if filename.endswith(".docx"):
                content = read_docx_full(filepath)
            elif filename.endswith(".xlsx"):
                content = read_xlsx_full(filepath)
            else:
                continue

            works = extract_meaningful_work(content)
            persons[person_dir]["works"].extend(works)

    for person_name, data in persons.items():
        if person_name == "示例员工":
            data["highlights"] = [
                "合同评审制度库：完成合同评审法律法规及公司制度的文档整理与分类",
                "客户拜访系统：完成销售客户拜访日志打分功能设计与实现",
                "资产管理系统：完成资产数据与U8系统对照，优化多维度报表功能",
                "无锡出差：完成派克新材料现场对接，了解AI助手应用场景",
            ]
        elif person_name == "水彬":
            data["highlights"] = [
                "WMS系统调研：完成现有系统业务流程和操作的全面熟悉",
                "数据库分析：完成新旧WMS数据库结构对照和单据流程梳理",
                "文档整理：协助整理WMS开发文档及设计文档",
                "技术学习：熟悉AI辅助编程工具，提升开发效率",
            ]

    html = generate_clean_summary_html(dinghai_data, shuibin_data)

    output_path = os.path.join(base_dir, "管理者汇总报告_20260518-20260526.html")
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"✅ 管理者汇总报告已生成: {output_path}")

if __name__ == "__main__":
    main()
