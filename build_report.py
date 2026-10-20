# -*- coding: utf-8 -*-
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from data import WEIGHTS, CANDIDATES, WAIT, REJECTED, POLICIES

OUT = "artifacts"
os.makedirs(OUT, exist_ok=True)

DARK = PatternFill("solid", fgColor="1F4E79")
GREEN = PatternFill("solid", fgColor="C6EFCE")
YELLOW = PatternFill("solid", fgColor="FFEB9C")
GRAY = PatternFill("solid", fgColor="E7E6E6")
WHITE_F = Font(color="FFFFFF", bold=True, size=11)
BOLD = Font(bold=True)
WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
THIN = Border(*[Side(style="thin", color="BFBFBF")] * 4)

wb = Workbook()

def header(ws, headers, widths):
    for i, h in enumerate(headers, 1):
        c = ws.cell(1, i, h)
        c.fill = DARK; c.font = WHITE_F; c.alignment = CENTER; c.border = THIN
        ws.column_dimensions[get_column_letter(i)].width = widths[i - 1]
    ws.freeze_panes = "A2"

# 1. 重排榜单
ws = wb.active; ws.title = "重排榜单"
hdr = ["排名", "ID", "机会", "目标人群", "形态", "总分", "分区", "一句话结论"]
header(ws, hdr, [6, 9, 28, 22, 26, 7, 10, 50])
rows = [(i + 1, c, GREEN, "高确定性") for i, c in enumerate(CANDIDATES)]
rows += [(len(CANDIDATES) + i + 1, w, YELLOW, "候补") for i, w in enumerate(WAIT)]
for r, (rank, item, fill, zone) in enumerate(rows, 2):
    name = item["name"]; sc = item["score"]
    onel = item.get("monetization", item.get("note", ""))
    vals = [rank, item["id"], name, item.get("audience", "-"), item.get("form", "-"), sc, zone, onel]
    for i, v in enumerate(vals, 1):
        c = ws.cell(r, i, v); c.border = THIN; c.alignment = WRAP if i in (3, 4, 5, 8) else CENTER
        if i in (6, 7): c.fill = fill
    ws.row_dimensions[r].height = 42

# 2. 评分明细
ws = wb.create_sheet("评分明细")
dim_names = list(WEIGHTS.keys())
hdr = ["ID", "机会", "总分"] + [f"{k}({v})" for k, v in WEIGHTS.items()]
header(ws, hdr, [9, 28, 7] + [11] * 7)
for r, c in enumerate(CANDIDATES, 2):
    ws.cell(r, 1, c["id"]).border = THIN
    ws.cell(r, 2, c["name"]).border = THIN
    ws.cell(r, 3, c["score"]).border = THIN
    d = c["dim"]
    keys = ["付费", "痛点", "竞品", "频次", "付费意愿", "免费替代", "薄壳"]
    for i, k in enumerate(keys):
        v = "-" if d is None else d.get(k, "-")
        cell = ws.cell(r, 4 + i, v); cell.border = THIN; cell.alignment = CENTER
    for i in (1, 2, 3):
        ws.cell(r, i).alignment = WRAP if i == 2 else CENTER
ws.cell(len(CANDIDATES) + 3, 1, "注：仅D4流回报完整七维分；其余流为总分+证据，明细不补造").font = Font(italic=True, color="808080")

# 3. 证据档案
ws = wb.create_sheet("证据档案")
hdr = ["ID", "机会", "字段", "内容"]
header(ws, hdr, [9, 26, 16, 100])
r = 2
for c in CANDIDATES:
    blocks = [
        ("目标人群", c["audience"]), ("形态", c["form"]),
        ("变现方式", c["monetization"]), ("冷启动机制（具体步骤）", c["cold_start"]),
        ("平台为什么不做", c["why_not_platform"]), ("成本结构", c["cost"]),
        ("30天里程碑与止损", c["milestone"]), ("风险", c["risk"]),
    ]
    for q, e in c["evidence"]:
        blocks.append(("证据：" + q, e))
    start = r
    for k, v in blocks:
        ws.cell(r, 3, k).alignment = WRAP; ws.cell(r, 3).border = THIN
        cell = ws.cell(r, 4, v); cell.alignment = WRAP; cell.border = THIN
        if k.startswith("证据"): ws.cell(r, 3).font = Font(color="C00000")
        r += 1
    ws.cell(start, 1, c["id"])
    ws.cell(start, 2, f"{c['name']}（{c['score']}分）")
    ws.merge_cells(start_row=start, start_column=1, end_row=r - 1, end_column=1)
    ws.merge_cells(start_row=start, start_column=2, end_row=r - 1, end_column=2)
    ws.cell(start, 1).alignment = CENTER; ws.cell(start, 2).alignment = CENTER

# 4. 候补区
ws = wb.create_sheet("候补区")
header(ws, ["ID", "机会", "分数", "说明/升级条件"], [9, 30, 7, 90])
for r, w in enumerate(WAIT, 2):
    vals = [w["id"], w["name"], w["score"], w["note"]]
    for i, v in enumerate(vals, 1):
        cell = ws.cell(r, i, v); cell.border = THIN; cell.alignment = WRAP if i in (2, 4) else CENTER
        cell.fill = YELLOW if i == 3 else PatternFill()

# 5. 证伪台账
ws = wb.create_sheet("证伪台账")
header(ws, ["方向", "证伪原因", "来源"], [34, 90, 40])
for r, (name, why, url) in enumerate(REJECTED, 2):
    ws.cell(r, 1, name).alignment = WRAP; ws.cell(r, 1).border = THIN
    ws.cell(r, 2, why).alignment = WRAP; ws.cell(r, 2).border = THIN
    cell = ws.cell(r, 3, url); cell.border = THIN; cell.hyperlink = url if url else None
    cell.font = Font(color="0563C1", underline="single") if url else Font()

# 6. 平台政策
ws = wb.create_sheet("平台政策事实")
header(ws, ["事项", "事实（含口径冲突）", "来源"], [20, 110, 30])
for r, (k, v, url) in enumerate(POLICIES, 2):
    ws.cell(r, 1, k).alignment = WRAP; ws.cell(r, 1).border = THIN; ws.cell(r, 1).font = BOLD
    ws.cell(r, 2, v).alignment = WRAP; ws.cell(r, 2).border = THIN
    ws.cell(r, 3, url).border = THIN

# 7. 结论
ws = wb.create_sheet("结论")
header(ws, ["要点", "内容"], [24, 110])
concl = [
 ("最终通过", f"{len(CANDIDATES)}个高确定性方向（60-73分）：怀旧老歌公众号、心理测试付费报告、老照片修复、长尾工具小程序矩阵、蓝领记工、AI尽孝、随礼记账、AI修图交付"),
 ("核心判断", "下沉C端的真实钱在‘自然流量+情绪/仪式感+人工交付’，不在投流；平台不做的是脏内容/脏数据/重服务环节"),
 ("优先顺序", "1)怀旧老歌公众号（轻、快、eCPM高）→ 2)老照片修复（有政府报道案例+跨境溢价）→ 3)长尾工具矩阵（流量主，需个体户）→ 4)心理测试（付费转化）"),
 ("合规红线", "算命/塔罗/风水收费、诱导分享、短剧小说CPS、AI不标识，均不碰"),
 ("前置动作", "办个体工商户（300元/年）解锁流量主与虚拟支付；政策细节以官方文档为准"),
]
for r, (k, v) in enumerate(concl, 2):
    ws.cell(r, 1, k).alignment = WRAP; ws.cell(r, 1).font = BOLD; ws.cell(r, 1).border = THIN
    ws.cell(r, 2, v).alignment = WRAP; ws.cell(r, 2).border = THIN

xlsx_path = os.path.join(OUT, "C端下沉市场机会报告_2026Q4.xlsx")
wb.save(xlsx_path)
print("saved:", xlsx_path)

# ---- Markdown ----
md = []
md.append("# C端下沉市场机会报告（2026 Q4）\n")
md.append("> 方法：4 路扫描（中老年/宝妈青年蓝领/工具变现/玄学合规），逐个七维实证，证据不足不入选。\n")
md.append("## 一、重排榜单\n")
md.append("| 排名 | ID | 机会 | 人群 | 形态 | 总分 | 分区 |")
md.append("|---|---|---|---|---|---|---|")
for i, c in enumerate(CANDIDATES, 1):
    md.append(f"| {i} | {c['id']} | {c['name']} | {c['audience']} | {c['form']} | **{c['score']}** | 🟢高确定性 |")
for i, w in enumerate(WAIT, len(CANDIDATES) + 1):
    md.append(f"| {i} | {w['id']} | {w['name']} | - | - | {w['score']} | 🟡候补 |")
md.append("\n## 二、高确定性机会明细\n")
for c in CANDIDATES:
    md.append(f"### {c['id']} {c['name']}（{c['score']}分）")
    md.append(f"- **人群/形态**：{c['audience']}｜{c['form']}")
    md.append(f"- **变现**：{c['monetization']}")
    md.append(f"- **冷启动（具体步骤）**：{c['cold_start']}")
    md.append(f"- **平台为什么不做**：{c['why_not_platform']}")
    md.append(f"- **成本**：{c['cost']}")
    md.append(f"- **里程碑与止损**：{c['milestone']}")
    md.append(f"- **风险**：{c['risk']}")
    for q, e in c["evidence"]:
        md.append(f"  - 证据：{q} — {e}")
    md.append("")
md.append("## 三、候补区\n")
for w in WAIT:
    md.append(f"- **{w['name']}（{w['score']}）**：{w['note']} {w['evidence']}")
md.append("\n## 四、证伪台账\n")
for name, why, url in REJECTED:
    md.append(f"- **{name}**：{why} {url}")
md.append("\n## 五、平台政策事实\n")
for k, v, url in POLICIES:
    md.append(f"- **{k}**：{v}")
md.append("")
md_path = os.path.join(OUT, "C端下沉市场机会报告_2026Q4.md")
with open(md_path, "w", encoding="utf-8") as f:
    f.write("\n".join(md))
print("saved:", md_path)

# ---- 回读自检 ----
from openpyxl import load_workbook
wb2 = load_workbook(xlsx_path)
print("sheets:", wb2.sheetnames)
for s in wb2.sheetnames:
    print(s, wb2[s].max_row, wb2[s].max_column)
