"""单元格依赖图：构图、循环引用检测、影响面/依赖链查询。"""
from __future__ import annotations

from collections import deque
from dataclasses import dataclass, field

from openpyxl.utils import get_column_letter

from .formula_parser import RawRef, parse_formula

# 单个区域最多展开的单元格数，防止 SUM(A:XFD) 之类打爆内存
MAX_RANGE_CELLS = 5000
# 影响面查询的默认上限
MAX_IMPACT_NODES = 5000


def node_id(sheet: str, coord: str) -> str:
    return f"{sheet}!{coord}"


def split_node(node: str) -> tuple[str, str]:
    """工作表名本身可含 '!'，坐标部分不会，故从右侧切分。"""
    sheet, coord = node.rsplit("!", 1)
    return sheet, coord


def display_id(sheet: str, coord: str) -> str:
    """展示用 id：含特殊字符的工作表名加引号。"""
    if sheet.replace("_", "").isalnum():
        return f"{sheet}!{coord}"
    return f"'{sheet}'!{coord}"


@dataclass
class SheetInfo:
    name: str
    max_row: int
    max_col: int
    formula_count: int = 0


@dataclass
class WorkbookGraph:
    sheets: list[SheetInfo] = field(default_factory=list)
    formulas: dict[str, str] = field(default_factory=dict)      # node -> 公式文本
    values: dict[str, str] = field(default_factory=dict)        # node -> 缓存计算值
    precedents: dict[str, set[str]] = field(default_factory=dict)  # node -> 它引用的节点
    dependents: dict[str, set[str]] = field(default_factory=dict)  # node -> 引用它的节点
    cycles: list[list[str]] = field(default_factory=list)
    cycle_nodes: set[str] = field(default_factory=set)
    edge_count: int = 0

    # ------------------------------------------------------------------ 构图
    @classmethod
    def from_workbooks(cls, wb_formula, wb_values) -> "WorkbookGraph":
        g = cls()
        bounds: dict[str, tuple[int, int]] = {}
        for ws in wb_formula.worksheets:
            bounds[ws.title] = (ws.max_row or 1, ws.max_column or 1)
            g.sheets.append(SheetInfo(ws.title, ws.max_row or 1, ws.max_column or 1))

        ws_values = {ws.title: ws for ws in wb_values.worksheets}
        sheet_names = set(bounds)

        for ws in wb_formula.worksheets:
            wsv = ws_values.get(ws.title)
            info = next(s for s in g.sheets if s.name == ws.title)
            for row in ws.iter_rows():
                for cell in row:
                    formula = cell.value
                    if not isinstance(formula, str) or not formula.startswith("="):
                        continue
                    coord = cell.coordinate
                    node = node_id(ws.title, coord)
                    g.formulas[node] = formula
                    info.formula_count += 1
                    if wsv is not None:
                        v = wsv[coord].value
                        if v is not None:
                            g.values[node] = str(v)
                    for ref in parse_formula(formula):
                        target_sheet = ref.sheet if ref.sheet is not None else ws.title
                        if target_sheet not in sheet_names:
                            continue  # 引用不存在的工作表，跳过
                        wst = ws_values.get(target_sheet)
                        for tcoord in g._expand(ref, bounds[target_sheet]):
                            tnode = node_id(target_sheet, tcoord)
                            g.precedents.setdefault(node, set()).add(tnode)
                            g.dependents.setdefault(tnode, set()).add(node)
                            g.edge_count += 1
                            # 叶子单元格也记录值，便于前端展示
                            if tnode not in g.formulas and tnode not in g.values and wst is not None:
                                v = wst[tcoord].value
                                if v is not None:
                                    g.values[tnode] = str(v)

        g._detect_cycles()
        return g

    @staticmethod
    def _expand(ref: RawRef, bounds: tuple[int, int]) -> list[str]:
        """把 RawRef 展开为坐标列表。

        显式区域（A1:B10）按原样展开；整行/整列区域按工作表已用范围截断；
        超过 MAX_RANGE_CELLS 时等比收缩，防止内存被打爆。
        """
        max_row, max_col = bounds
        if ref.kind == "cell":
            return [f"{get_column_letter(ref.col1)}{ref.row1}"]
        if ref.kind == "range":
            r1, r2, c1, c2 = ref.row1, ref.row2, ref.col1, ref.col2
        elif ref.kind == "cols":
            r1, r2 = 1, max_row
            c1, c2 = ref.col1, min(ref.col2, max_col)
        else:  # rows
            r1, r2 = ref.row1, ref.row2
            c1, c2 = 1, max_col
        while (r2 - r1 + 1) * (c2 - c1 + 1) > MAX_RANGE_CELLS and r2 > r1:
            r2 -= 1
        while (r2 - r1 + 1) * (c2 - c1 + 1) > MAX_RANGE_CELLS and c2 > c1:
            c2 -= 1
        return [
            f"{get_column_letter(c)}{r}"
            for r in range(r1, r2 + 1)
            for c in range(c1, c2 + 1)
        ]

    # -------------------------------------------------------------- 环检测
    def _detect_cycles(self) -> None:
        """Tarjan 强连通分量（迭代实现）。大小>1 的分量或自环即循环引用。"""
        adj = self.precedents
        index_of: dict[str, int] = {}
        lowlink: dict[str, int] = {}
        on_stack: set[str] = set()
        stack: list[str] = []
        counter = 0

        for root in list(adj):
            if root in index_of:
                continue
            work = [(root, iter(adj.get(root, ())))]
            index_of[root] = lowlink[root] = counter
            counter += 1
            stack.append(root)
            on_stack.add(root)
            while work:
                node, it = work[-1]
                advanced = False
                for nxt in it:
                    if nxt not in adj and nxt not in index_of:
                        continue  # 非公式叶子节点不参与环
                    if nxt not in index_of:
                        index_of[nxt] = lowlink[nxt] = counter
                        counter += 1
                        stack.append(nxt)
                        on_stack.add(nxt)
                        work.append((nxt, iter(adj.get(nxt, ()))))
                        advanced = True
                        break
                    elif nxt in on_stack:
                        lowlink[node] = min(lowlink[node], index_of[nxt])
                if advanced:
                    continue
                work.pop()
                if work:
                    parent = work[-1][0]
                    lowlink[parent] = min(lowlink[parent], lowlink[node])
                if lowlink[node] == index_of[node]:
                    scc = []
                    while True:
                        w = stack.pop()
                        on_stack.discard(w)
                        scc.append(w)
                        if w == node:
                            break
                    is_cycle = len(scc) > 1 or any(
                        n in adj.get(n, ()) for n in scc
                    )
                    if is_cycle:
                        self.cycles.append(sorted(scc))
                        self.cycle_nodes.update(scc)

    # -------------------------------------------------------------- 查询
    def neighbors(self, node: str, direction: str) -> set[str]:
        adj = self.dependents if direction == "downstream" else self.precedents
        return adj.get(node, set())

    def impact(
        self,
        node: str,
        direction: str = "downstream",
        max_depth: int = 50,
        max_nodes: int = MAX_IMPACT_NODES,
    ) -> dict:
        """BFS 求某单元格的影响面（下游）或依赖链（上游）。

        返回每个可达节点的最短路径；被环截断的节点会标记 in_cycle。
        """
        adj = self.dependents if direction == "downstream" else self.precedents
        depth_of = {node: 0}
        parent = {node: None}
        truncated = False
        queue = deque([node])
        while queue:
            cur = queue.popleft()
            d = depth_of[cur]
            if d >= max_depth:
                continue
            for nxt in sorted(adj.get(cur, ())):
                if nxt in depth_of:
                    continue
                if len(depth_of) >= max_nodes:
                    truncated = True
                    break
                depth_of[nxt] = d + 1
                parent[nxt] = cur
                queue.append(nxt)
            if truncated:
                break

        def path_to(n: str) -> list[str]:
            path = []
            while n is not None:
                path.append(n)
                n = parent[n]
            return list(reversed(path))

        items = [
            {
                "cell": n,
                "sheet": split_node(n)[0],
                "coord": split_node(n)[1],
                "display": display_id(*split_node(n)),
                "depth": depth_of[n],
                "formula": self.formulas.get(n),
                "value": self.values.get(n),
                "in_cycle": n in self.cycle_nodes,
                "path": [display_id(*split_node(p)) for p in path_to(n)],
                "path_raw": path_to(n),
            }
            for n in depth_of
            if n != node
        ]
        items.sort(key=lambda x: (x["depth"], x["cell"]))
        direct = sorted(adj.get(node, set()))
        return {
            "root": node,
            "direction": direction,
            "direct": [
                {
                    "cell": n,
                    "display": display_id(*split_node(n)),
                    "formula": self.formulas.get(n),
                    "value": self.values.get(n),
                    "in_cycle": n in self.cycle_nodes,
                }
                for n in direct
            ],
            "impacted": items,
            "truncated": truncated,
            "total": len(items),
        }

    # -------------------------------------------------------------- 汇总
    def summary(self) -> dict:
        return {
            "sheets": [
                {
                    "name": s.name,
                    "max_row": s.max_row,
                    "max_col": s.max_col,
                    "formula_count": s.formula_count,
                }
                for s in self.sheets
            ],
            "formula_count": len(self.formulas),
            "edge_count": self.edge_count,
            "cycle_count": len(self.cycles),
            "cycles": [
                [display_id(*split_node(n)) for n in cyc] for cyc in self.cycles
            ],
        }

    def cell_list(
        self,
        sheet: str | None = None,
        query: str = "",
        formulas_only: bool = True,
        limit: int = 1000,
        offset: int = 0,
    ) -> dict:
        nodes = set(self.formulas)
        if not formulas_only:
            nodes |= set(self.dependents)  # 被引用的非公式单元格
        rows = []
        q = query.lower()
        for n in sorted(nodes):
            s, coord = split_node(n)
            if sheet and s != sheet:
                continue
            disp = display_id(s, coord)
            if q and q not in disp.lower() and q not in (self.formulas.get(n) or "").lower():
                continue
            rows.append(
                {
                    "cell": n,
                    "sheet": s,
                    "coord": coord,
                    "display": disp,
                    "formula": self.formulas.get(n),
                    "value": self.values.get(n),
                    "in_cycle": n in self.cycle_nodes,
                    "dependent_count": len(self.dependents.get(n, ())),
                    "precedent_count": len(self.precedents.get(n, ())),
                }
            )
        return {"total": len(rows), "cells": rows[offset : offset + limit]}

    def export(self, filename: str) -> dict:
        """完整分析结果，供导出 JSON。"""
        edges = [
            {"from": src, "to": dst}
            for src, dsts in sorted(self.dependents.items())
            for dst in sorted(dsts)
        ]
        return {
            "filename": filename,
            "summary": self.summary(),
            "formulas": {
                display_id(*split_node(n)): {
                    "formula": f,
                    "value": self.values.get(n),
                    "in_cycle": n in self.cycle_nodes,
                }
                for n, f in sorted(self.formulas.items())
            },
            "edges": [
                {"from": display_id(*split_node(e["from"])), "to": display_id(*split_node(e["to"]))}
                for e in edges
            ],
            "cycles": [
                [display_id(*split_node(n)) for n in cyc] for cyc in self.cycles
            ],
        }
