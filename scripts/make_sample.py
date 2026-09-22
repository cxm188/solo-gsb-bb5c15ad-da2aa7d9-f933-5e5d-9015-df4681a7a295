"""生成示例工作簿 samples/sample_impact.xlsx。

覆盖场景：
  - 跨工作表引用（=参数!B1*参数!B2）
  - 带引号/空格的工作表名（='月报 2026'!...）
  - 矩形区域引用（=SUM(计算!B1:B4)）
  - 整列引用（=SUM(参数!B:B)）
  - 多级依赖链（参数 -> 计算 -> 汇总 -> 月报）
  - 循环引用（循环!A1 <-> 循环!B1，以及 循环!C1 自环）
"""
from pathlib import Path

from openpyxl import Workbook

OUT = Path(__file__).resolve().parent.parent / "samples" / "sample_impact.xlsx"


def main() -> None:
    wb = Workbook()

    # ---- 参数表：纯输入 ----
    ws = wb.active
    ws.title = "参数"
    rows = [
        ("项目", "值"),
        ("单价", 10),
        ("数量", 5),
        ("折扣率", 0.9),
        ("税率", 0.13),
        ("月数", 12),
    ]
    for r in rows:
        ws.append(r)

    # ---- 计算表：跨表引用 + 链式依赖 ----
    ws = wb.create_sheet("计算")
    ws.append(["项目", "值"])
    ws["A2"] = "小计"
    ws["B2"] = "=参数!B2*参数!B3"          # 单价 * 数量
    ws["A3"] = "折后"
    ws["B3"] = "=B2*参数!B4"               # 小计 * 折扣率
    ws["A4"] = "税额"
    ws["B4"] = "=B3*参数!B5"               # 折后 * 税率
    ws["A5"] = "月合计"
    ws["B5"] = "=B3+B4"                    # 折后 + 税额
    ws["A6"] = "年合计"
    ws["B6"] = "=B5*参数!B6"               # 月合计 * 月数

    # ---- 汇总表：跨表矩形区域引用 ----
    ws = wb.create_sheet("汇总")
    ws.append(["指标", "值"])
    ws["A2"] = "年合计(跨表)"
    ws["B2"] = "=计算!B6"
    ws["A3"] = "季度合计"
    ws["B3"] = "=计算!B6/4"
    ws["A4"] = "计算区求和(区域)"
    ws["B4"] = "=SUM(计算!B2:B5)"
    ws["A5"] = "参数区求和(区域)"
    ws["B5"] = "=SUM(参数!B2:B6)"
    ws["A6"] = "双区合计"
    ws["B6"] = "=SUM(计算!B2:B5)+SUM(参数!B2:B5)"
    ws["A7"] = "整列引用"
    ws["B7"] = "=SUM(参数!B:B)"

    # ---- 月报 2026：带空格的工作表名，公式内需引号 ----
    ws = wb.create_sheet("月报 2026")
    ws.append(["月份", "月均"])
    ws["A2"] = "月均(来自年合计)"
    ws["B2"] = "=计算!B6/12"
    ws["A3"] = "月均(来自汇总)"
    ws["B3"] = "=汇总!B2/12"
    ws["A4"] = "校验"
    ws["B4"] = "=B2-B3"

    # ---- 循环表：故意制造循环引用，用于演示检测 ----
    ws = wb.create_sheet("循环")
    ws["A1"] = "=B1+1"
    ws["B1"] = "=A1+1"
    ws["C1"] = "=C1*2"                     # 自环
    ws["D1"] = "=A1+10"                    # 受循环影响的下游单元格

    OUT.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUT)
    print(f"已生成 {OUT}")


if __name__ == "__main__":
    main()
