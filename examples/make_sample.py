"""生成示例工作簿 examples/sample_formulas.xlsx。

包含场景：
  * 同表与跨工作表引用
  * 矩形区域引用（SUM(A1:A10) 等）与整列引用
  * 多级间接影响链（修改基础数据后影响汇总 / 报告表）
  * 定义名称
  * 结构化表引用
  * 循环引用（在独立区域，便于观察检测结果）

运行：python examples/make_sample.py
"""
import os

from openpyxl import Workbook

OUT = os.path.join(os.path.dirname(__file__), "sample_formulas.xlsx")


def build() -> Workbook:
    wb = Workbook()

    # ---------------- 基础数据 ---------------- #
    data = wb.active
    data.title = "基础数据"
    data["A1"] = "月份"
    data["B1"] = "销售额"
    data["C1"] = "成本"
    rows = [
        ("1月", 1200, 700),
        ("2月", 1500, 900),
        ("3月", 1800, 950),
        ("4月", 1100, 680),
        ("5月", 2100, 1200),
        ("6月", 1900, 1050),
    ]
    for i, (m, s, c) in enumerate(rows, start=2):
        data.cell(row=i, column=1, value=m)
        data.cell(row=i, column=2, value=s)
        data.cell(row=i, column=3, value=c)
    # 矩形区域引用
    data["E1"] = "总销售额"
    data["F1"] = "=SUM(B2:B7)"
    data["E2"] = "总成本"
    data["F2"] = "=SUM(C2:C7)"
    data["E3"] = "总利润"
    data["F3"] = "=F1-F2"
    data["E4"] = "利润率"
    data["F4"] = "=F3/F1"
    # 整列引用（展开时受工作表维度限制）
    data["E6"] = "平均销售额（整列）"
    data["F6"] = "=AVERAGE(B:B)"

    # ---------------- 汇总（跨表引用基础数据） ---------------- #
    summary = wb.create_sheet("季度汇总")
    summary["A1"] = "指标"
    summary["B1"] = "数值"
    summary["A2"] = "上半年销售额"
    summary["B2"] = "='基础数据'!F1"
    summary["A3"] = "上半年成本"
    summary["B3"] = "='基础数据'!F2"
    summary["A4"] = "上半年利润"
    summary["B4"] = "=B2-B3"
    summary["A5"] = "利润率"
    summary["B5"] = "=B4/B2"
    # 定义名称：税率
    from openpyxl.workbook.defined_name import DefinedName

    wb.defined_names["税率"] = DefinedName(name="税率", attr_text="0.13")
    summary["A6"] = "税后利润"
    summary["B6"] = "=B4*(1-税率)"

    # ---------------- 报告（引用汇总表，形成多级影响链） ---------------- #
    report = wb.create_sheet("管理层报告")
    report["A1"] = "关键结论"
    report["B1"] = "结果"
    report["A2"] = "半年营收"
    report["B2"] = "=季度汇总!B2"
    report["A3"] = "半年净利润"
    report["B3"] = "=季度汇总!B6"
    report["A4"] = "结论"
    report["B4"] = '=IF(B3>3000,"盈利良好","需要关注")'
    report["A5"] = "引用报告自身历史结论的校验格"
    report["B5"] = "=季度汇总!B5"

    # ---------------- 结构化表 ---------------- #
    tbl_ws = wb.create_sheet("产品表")
    tbl_ws.append(["产品", "单价", "数量"])
    tbl_ws.append(["键盘", 199, 30])
    tbl_ws.append(["鼠标", 89, 120])
    tbl_ws.append(["显示器", 1299, 12])
    from openpyxl.worksheet.table import Table, TableStyleInfo

    table = Table(displayName="产品明细", ref="A1:C4")
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium9", showRowStripes=True
    )
    tbl_ws.add_table(table)
    tbl_ws["E1"] = "库存总值（结构化表区域）"
    tbl_ws["F1"] = "=SUMPRODUCT(产品明细[单价],产品明细[数量])"

    # ---------------- 循环引用演示 ---------------- #
    cyc = wb.create_sheet("循环引用演示")
    cyc["A1"] = "以下三个单元格构成 A3 -> B3 -> C3 -> A3 的环："
    cyc["A3"] = "=B3+1"
    cyc["B3"] = "=C3+1"
    cyc["C3"] = "=A3+1"
    cyc["A6"] = "自引用循环："
    cyc["B6"] = "=B6+1"

    return wb


def main() -> None:
    wb = build()
    wb.save(OUT)
    print(f"已生成示例文件: {OUT}")


if __name__ == "__main__":
    main()
