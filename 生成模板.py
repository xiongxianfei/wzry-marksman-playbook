"""
运行此脚本生成 Excel 记录模板：
    python 生成模板.py
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

OUTPUT = "训练记录.xlsx"

# ── 颜色 ──────────────────────────────────────────────
C_HEADER   = "1F2937"   # 深灰（表头背景）
C_SUBHEAD  = "374151"   # 次级表头
C_GOLD     = "F0C040"   # 金色（标题字）
C_WHITE    = "FFFFFF"
C_ALT      = "F9FAFB"   # 隔行浅灰
C_GREEN    = "D1FAE5"
C_RED      = "FEE2E2"

def header_fill(color): return PatternFill("solid", fgColor=color)
def font(bold=False, color=C_WHITE, size=11):
    return Font(bold=bold, color=color, name="微软雅黑", size=size)
def center(): return Alignment(horizontal="center", vertical="center")
def left():   return Alignment(horizontal="left",   vertical="center")
def border():
    s = Side(style="thin", color="D1D5DB")
    return Border(left=s, right=s, top=s, bottom=s)


def make_record_sheet(wb):
    ws = wb.active
    ws.title = "每局记录"
    ws.sheet_view.showGridLines = False

    # ── 列定义 ──────────────────────────────────────────
    cols = [
        ("日期",       12),
        ("英雄",       10),
        ("胜负",        8),
        ("10分钟经济",  14),
        ("死亡次数",    10),
        ("暴君参与",    10),
        ("决策三问",    10),
        ("9分靠队伍",   10),
        ("推塔",         8),
        ("跟强打",       8),
        ("心态稳定",    10),
        ("复盘分",      10),
        ("备注",        30),
    ]

    # 列宽
    for i, (_, w) in enumerate(cols, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # ── 标题行 ──────────────────────────────────────────
    ws.row_dimensions[1].height = 32
    ws.merge_cells("A1:M1")
    c = ws["A1"]
    c.value          = "王者荣耀发育路训练记录"
    c.font           = Font(bold=True, color=C_GOLD, name="微软雅黑", size=14)
    c.fill           = header_fill(C_HEADER)
    c.alignment      = center()

    # ── 表头行 ──────────────────────────────────────────
    ws.row_dimensions[2].height = 24
    for i, (name, _) in enumerate(cols, 1):
        c = ws.cell(row=2, column=i, value=name)
        c.font      = font(bold=True, color=C_GOLD)
        c.fill      = header_fill(C_SUBHEAD)
        c.alignment = center()
        c.border    = border()

    # ── 数据验证 ──────────────────────────────────────────
    heroes = '"后羿,莱西奥,艾琳,戈娅,孙尚香"'
    yesno  = '"是,否"'
    winlos = '"胜,负"'

    def dv(formula, cols_range):
        v = DataValidation(type="list", formula1=formula, allow_blank=True)
        v.error      = "请从列表选择"
        v.errorTitle = "输入错误"
        ws.add_data_validation(v)
        v.add(cols_range)

    dv(heroes, "B3:B1000")
    dv(winlos, "C3:C1000")
    dv(yesno,  "F3:F1000")
    dv(yesno,  "G3:G1000")
    dv(yesno,  "H3:H1000")
    dv(yesno,  "I3:I1000")
    dv(yesno,  "J3:J1000")
    dv(yesno,  "K3:K1000")

    # ── 复盘分公式（满分100）──────────────────────────────
    # 经济达标(15) + 死亡达标(10) + 暴君(10) + 三问(10) + 靠队伍(10)
    # + 推塔(10) + 跟强打(10) + 心态(15) = 90；备注填写+10
    def score_formula(row):
        return (
            f'=IF(D{row}>=6500,15,0)'
            f'+IF(E{row}<=2,10,0)'
            f'+IF(F{row}="是",10,0)'
            f'+IF(G{row}="是",10,0)'
            f'+IF(H{row}="是",10,0)'
            f'+IF(I{row}="是",10,0)'
            f'+IF(J{row}="是",10,0)'
            f'+IF(K{row}="是",15,0)'
            f'+IF(M{row}<>"",10,0)'
        )

    # ── 50行数据区 ──────────────────────────────────────
    for row in range(3, 103):
        ws.row_dimensions[row].height = 20
        fill = PatternFill("solid", fgColor=C_ALT if row % 2 == 0 else C_WHITE)
        for col in range(1, 14):
            c = ws.cell(row=row, column=col)
            c.fill      = fill
            c.border    = border()
            c.alignment = center() if col != 13 else left()
            c.font      = Font(name="微软雅黑", size=10, color="374151")
        # 复盘分公式
        ws.cell(row=row, column=12).value = score_formula(row)

    # ── 条件格式提示（用公式着色） ──────────────────────
    from openpyxl.formatting.rule import CellIsRule, FormulaRule
    # 经济列：达标绿，未达标红
    red_fill   = PatternFill("solid", fgColor="FEE2E2")
    green_fill = PatternFill("solid", fgColor="D1FAE5")
    ws.conditional_formatting.add(
        "D3:D102",
        CellIsRule(operator="greaterThanOrEqual", formula=["6500"], fill=green_fill)
    )
    ws.conditional_formatting.add(
        "D3:D102",
        CellIsRule(operator="lessThan", formula=["6500"], fill=red_fill)
    )
    # 死亡列：≤2绿，>2红
    ws.conditional_formatting.add(
        "E3:E102",
        CellIsRule(operator="lessThanOrEqual", formula=["2"], fill=green_fill)
    )
    ws.conditional_formatting.add(
        "E3:E102",
        CellIsRule(operator="greaterThan", formula=["2"], fill=red_fill)
    )
    # 复盘分：≥80绿，<60红
    ws.conditional_formatting.add(
        "L3:L102",
        CellIsRule(operator="greaterThanOrEqual", formula=["80"], fill=green_fill)
    )
    ws.conditional_formatting.add(
        "L3:L102",
        CellIsRule(operator="lessThan", formula=["60"], fill=red_fill)
    )

    # 冻结前两行
    ws.freeze_panes = "A3"


def make_stats_sheet(wb):
    ws = wb.create_sheet("周统计")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 16

    def title(row, text):
        ws.row_dimensions[row].height = 26
        c = ws.cell(row=row, column=1, value=text)
        c.font      = Font(bold=True, color=C_GOLD, name="微软雅黑", size=12)
        c.fill      = header_fill(C_HEADER)
        c.alignment = left()
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)

    def row_item(r, label, formula, note=""):
        ws.row_dimensions[r].height = 20
        a = ws.cell(row=r, column=1, value=label)
        a.font = Font(name="微软雅黑", size=10, color="374151")
        a.fill = PatternFill("solid", fgColor=C_ALT if r % 2 == 0 else C_WHITE)
        a.border = border()
        b = ws.cell(row=r, column=2, value=formula)
        b.font = Font(bold=True, name="微软雅黑", size=10, color="1F2937")
        b.fill = PatternFill("solid", fgColor=C_ALT if r % 2 == 0 else C_WHITE)
        b.border = border()
        b.alignment = center()
        c = ws.cell(row=r, column=3, value=note)
        c.font = Font(name="微软雅黑", size=9, color="9CA3AF", italic=True)
        c.fill = PatternFill("solid", fgColor=C_ALT if r % 2 == 0 else C_WHITE)
        c.border = border()

    title(1, "全局统计（所有记录）")
    row_item(2,  "总场次",       "=COUNTA(每局记录!A3:A102)",                "局")
    row_item(3,  "总胜场",       '=COUNTIF(每局记录!C3:C102,"胜")',           "局")
    row_item(4,  "胜率",         '=IFERROR(COUNTIF(每局记录!C3:C102,"胜")/COUNTA(每局记录!C3:C102),"-")', "%")
    row_item(5,  "均10分钟经济", "=IFERROR(AVERAGEIF(每局记录!D3:D102,\">0\"),\"-\")",  "金币")
    row_item(6,  "均死亡次数",   "=IFERROR(AVERAGE(每局记录!E3:E102),\"-\")", "次")
    row_item(7,  "均复盘分",     "=IFERROR(AVERAGEIF(每局记录!L3:L102,\">0\"),\"-\")", "分")
    row_item(8,  "暴君参与率",   '=IFERROR(COUNTIF(每局记录!F3:F102,"是")/COUNTA(每局记录!F3:F102),"-")', "%")

    title(10, "各英雄场次分布")
    for i, hero in enumerate(["后羿", "莱西奥", "艾琳", "戈娅", "孙尚香"]):
        row_item(11+i, hero,
                 f'=COUNTIF(每局记录!B3:B102,"{hero}")',
                 f'胜: =COUNTIFS(每局记录!B3:B102,"{hero}",每局记录!C3:C102,"胜")')

    title(17, "目标达成率")
    row_item(18, "10分钟经济≥6500", '=IFERROR(COUNTIF(每局记录!D3:D102,">=6500")/COUNTA(每局记录!D3:D102),"-")', "%")
    row_item(19, "死亡≤2次",        '=IFERROR(COUNTIF(每局记录!E3:E102,"<=2")/COUNTA(每局记录!E3:E102),"-")',  "%")
    row_item(20, "复盘分≥80",       '=IFERROR(COUNTIF(每局记录!L3:L102,">=80")/COUNTA(每局记录!L3:L102),"-")', "%")
    row_item(21, "复盘分≥60",       '=IFERROR(COUNTIF(每局记录!L3:L102,">=60")/COUNTA(每局记录!L3:L102),"-")', "%")

    ws.freeze_panes = "A2"


def main():
    wb = Workbook()
    make_record_sheet(wb)
    make_stats_sheet(wb)
    wb.save(OUTPUT)
    print(f"[OK] 已生成：{OUTPUT}")
    print("  Sheet1-每局记录：填写每局数据，复盘分自动计算")
    print("  Sheet2-周统计：自动汇总所有统计数据")


if __name__ == "__main__":
    main()
