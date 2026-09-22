"""工作簿解析与公式依赖图分析。

核心模型
--------
节点：单元格 (工作表, 坐标)，坐标统一为大写 A1 形式（不带 ``$``）。
边  ：若单元格 X 的公式引用了 Y，则存在一条有向边 X -> Y。
      因此：
      * 下游影响（dependents，受 X 影响）= 沿出边反向（reverse graph）可达；
      * 上游依赖（precedents，X 依赖谁）= 沿出边正向可达。

只依赖 openpyxl 的公式 tokenizer，不计算公式值。
"""

from __future__ import annotations

import re
from collections import deque
from dataclasses import dataclass, field
from typing import Dict, Iterable, List, Optional, Set, Tuple

from openpyxl.formula.tokenizer import Tokenizer
from openpyxl.utils import range_boundaries
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.worksheet import Worksheet

# Excel 工作表上限
MAX_ROW = 1048576
MAX_COL = 16384

# 单个区域引用展开的最大单元格数（如 A:A、1:1 会产生百万级节点）
MAX_EXPANSION = 100_000
# 单个公式产生的引用边上限
MAX_EDGES_PER_FORMULA = 200_000
# 整张图的边上限（防止恶意 / 异常文件爆内存）
MAX_TOTAL_EDGES = 2_000_000

CellRef = Tuple[str, str]  # (sheet_name, COORD)

# 纯单元格：A1、$A$1
_CELL_RE_PARTS = r"\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6}"
# 单元格区域：A1:B2（两端都允许 $）
_CELL_RANGE_RE = re.compile(rf"^{_CELL_RE_PARTS}:{_CELL_RE_PARTS}$")
_CELL_RE = re.compile(rf"^{_CELL_RE_PARTS}$")
# 整列：A、A:C
_COL_RE = re.compile(r"^\$?[A-Za-z]{1,3}(?::\$?[A-Za-z]{1,3})?$")
# 整行：1、1:10
_ROW_RE = re.compile(r"^\$?[0-9]{1,7}(?::\$?[0-9]{1,7})?$")
# 结构化表引用：Table1[#All]、Table1[列名]
_STRUCTURED_RE = re.compile(r"^[^\[\]!]+\[.+\]$", re.S)


def _norm_coord(coord: str) -> Optional[str]:
    """把 A1 / $a$1 规范化为大写无 ``$``；非法（越界）返回 None。"""
    coord = coord.replace("$", "").strip()
    m = _CELL_RE.match(coord)
    if not m:
        return None
    try:
        col_letters, row = coordinate_from_string(coord)
        col = column_index_from_string(col_letters)
    except Exception:
        return None
    if col < 1 or col > MAX_COL or row < 1 or row > MAX_ROW:
        return None
    return f"{col_letters.upper()}{row}"


def _col_index(letters: str) -> Optional[int]:
    letters = letters.replace("$", "").strip()
    if not letters or not letters.isalpha() or len(letters) > 3:
        return None
    col = column_index_from_string(letters)
    if not 1 <= col <= MAX_COL:
        return None
    return col


def _row_number(text: str) -> Optional[int]:
    text = text.replace("$", "").strip()
    if not text.isdigit():
        return None
    row = int(text)
    if not 1 <= row <= MAX_ROW:
        return None
    return row


@dataclass
class Node:
    id: int
    sheet: str
    coord: str
    value: Optional[str] = None       # 常量值（非公式）的可显示文本
    formula: Optional[str] = None     # 公式原文（含前导 =）

    @property
    def key(self) -> str:
        return f"{self.sheet}!{self.coord}"


@dataclass
class Cycle:
    nodes: List[str] = field(default_factory=list)  # node keys，首尾相同表示闭环


@dataclass
class Analysis:
    filename: str
    sheets: List[Dict]
    nodes: List[Node]
    # node key -> [target key ...]，仅公式边
    edges: Dict[str, List[str]]
    external_refs: List[str]
    defined_names: Dict[str, str]
    cycles: List[Cycle]
    stats: Dict
    warnings: List[str]

    # 运行期索引
    _key_to_id: Dict[str, int] = field(default_factory=dict)
    _out: Dict[int, Set[int]] = field(default_factory=dict)
    _in: Dict[int, Set[int]] = field(default_factory=dict)

    def to_dict(self) -> Dict:
        return {
            "filename": self.filename,
            "sheets": self.sheets,
            "nodes": [
                {
                    "id": n.id,
                    "key": n.key,
                    "sheet": n.sheet,
                    "coord": n.coord,
                    "value": n.value,
                    "formula": n.formula,
                }
                for n in self.nodes
            ],
            "edges": self.edges,
            "external_refs": self.external_refs,
            "defined_names": self.defined_names,
            "cycles": [c.nodes for c in self.cycles],
            "stats": self.stats,
            "warnings": self.warnings,
        }


# --------------------------------------------------------------------------- #
# 构建
# ---------------------------------------------------------------------------
class _Builder:
    def __init__(self, wb, filename: str):
        self.wb = wb
        self.filename = filename
        self.warnings: List[str] = []

        self.nodes: Dict[CellRef, Node] = {}
        self.edges: Dict[str, Set[str]] = {}
        self.external_refs: Set[str] = set()
        self.defined_names: Dict[str, str] = {}
        self.truncated_formulas = 0
        self.missing_sheets: Set[str] = set()
        self.edge_total = 0

        # 工作表名 -> 实际名（不区分大小写）
        self._sheet_lookup = {ws.title.lower(): ws.title for ws in wb.worksheets}
        # 结构化表名 -> 表对象
        self._tables: Dict[str, object] = {}
        self._table_sheets: Dict[str, str] = {}
        for ws in wb.worksheets:
            for raw_name in (getattr(ws, "tables", {}) or {}).keys():
                try:
                    tbl = ws.tables[raw_name]
                except Exception:
                    continue
                if isinstance(tbl, str) or not getattr(tbl, "ref", None):
                    continue
                # displayName 才是公式中使用的名称
                dn = getattr(tbl, "displayName", None) or raw_name
                self._tables.setdefault(dn.casefold(), tbl)
                self._table_sheets.setdefault(dn.casefold(), ws.title)

    def _warn(self, msg: str) -> None:
        if msg not in self.warnings:
            self.warnings.append(msg)

    def _resolve_sheet(self, name: str) -> str:
        """解析工作表名（去引号），未知工作表返回规范化后的原名。"""
        name = name.strip()
        if len(name) >= 2 and name.startswith("'") and name.endswith("'"):
            name = name[1:-1].replace("''", "'")
        actual = self._sheet_lookup.get(name.lower())
        if actual is not None:
            return actual
        if name not in self.missing_sheets:
            self.missing_sheets.add(name)
            self._warn(f"公式引用了工作簿中不存在的工作表: {name}")
        return name

    def _node(self, sheet: str, coord: str, **kw) -> Node:
        ref = (sheet, coord)
        node = self.nodes.get(ref)
        if node is None:
            node = Node(id=len(self.nodes), sheet=sheet, coord=coord, **kw)
            self.nodes[ref] = node
        else:
            for k, v in kw.items():
                if v is not None and getattr(node, k, None) is None:
                    setattr(node, k, v)
        return node

    # -- 定义名称 ---------------------------------------------------------- #
    def _collect_defined_names(self) -> None:
        wb = self.wb
        # workbook 级
        for name, dn in (getattr(wb, "defined_names", {}) or {}).items():
            dest = self._defined_name_value(dn)
            if dest:
                self.defined_names[name] = dest
        # worksheet 级
        for ws in wb.worksheets:
            for name, dn in (getattr(ws, "defined_names", {}) or {}).items():
                dest = self._defined_name_value(dn)
                key = f"{ws.title}!{name}"
                if dest:
                    self.defined_names[key] = dest

    @staticmethod
    def _defined_name_value(dn) -> Optional[str]:
        if isinstance(dn, DefinedName):
            val = dn.value
        elif isinstance(dn, str):
            val = dn
        else:
            val = getattr(dn, "value", None)
        if isinstance(val, str):
            return val.strip()
        return None

    def _resolve_defined_name(
        self, raw: str, current_sheet: str, depth: int
    ) -> List[CellRef]:
        """把定义名称解析成单元格引用。先查 sheet 限定名，再查工作簿级。"""
        if depth > 3:
            return []
        key = raw
        target: Optional[str] = self.defined_names.get(key)
        sheet_ctx = current_sheet
        if target is None and "!" in raw:
            prefix, rest = raw.rsplit("!", 1)
            target = self.defined_names.get(rest)
            sheet_ctx = self._resolve_sheet(prefix)
        if target is None:
            # 不区分大小写再试一次
            lut = {k.lower(): v for k, v in self.defined_names.items()}
            target = lut.get(raw.lower())
        if target is None or target.startswith("#"):
            return []
        if not target.startswith("="):
            target = "=" + target
        refs: List[CellRef] = []
        for t in Tokenizer(target).items:
            if t.type == "OPERAND" and t.subtype == "RANGE":
                refs.extend(
                    self._expand_operand(t.value, sheet_ctx, depth + 1)
                )
        return refs

    # -- 区域展开 ---------------------------------------------------------- #
    def _expand_rect(
        self, sheet: str, min_col: int, min_row: int, max_col: int, max_row: int
    ) -> List[CellRef]:
        out: List[CellRef] = []
        total = (max_col - min_col + 1) * (max_row - min_row + 1)
        if total > MAX_EXPANSION:
            self._warn(
                f"工作表 {sheet} 上的区域含 {total} 个单元格，"
                f"超过展开上限 {MAX_EXPANSION}，仅展开前一部分"
            )
        count = 0
        from openpyxl.utils import get_column_letter

        for r in range(min_row, max_row + 1):
            for c in range(min_col, max_col + 1):
                out.append((sheet, f"{get_column_letter(c)}{r}"))
                count += 1
                if count >= MAX_EXPANSION:
                    return out
        return out

    def _sheet_dimensions(self, sheet: str) -> Tuple[int, int]:
        ws = self.wb[self._sheet_lookup[sheet.lower()]] if sheet.lower() in self._sheet_lookup else None
        if ws is not None:
            dim = ws.calculate_dimension()
            try:
                min_col, min_row, max_col, max_row = range_boundaries(dim)
                return max(max_col, 1), max(max_row, 1)
            except Exception:
                pass
        return 50, 1000  # 引用了不存在的工作表时的兜底范围

    def _expand_operand(
        self, value: str, current_sheet: str, depth: int = 0
    ) -> List[CellRef]:
        """展开一个 RANGE token 为单元格引用列表。"""
        raw = value.strip()

        # 外部工作簿引用：[1]Sheet1!A1
        if raw.startswith("["):
            self.external_refs.add(raw)
            return []

        # 拆 sheet 前缀
        sheet = current_sheet
        rest = raw
        if "!" in raw:
            prefix, rest = raw.rsplit("!", 1)
            # 前缀里仍带 [n] 也算外部引用
            if "[" in prefix:
                self.external_refs.add(raw)
                return []
            sheet = self._resolve_sheet(prefix)

        rest = rest.strip()

        # Sheet1!#REF! 等失效引用
        if rest.startswith("#"):
            return []

        # 单元格区域 A1:B2
        m = _CELL_RANGE_RE.match(rest)
        if m and ":" in rest:
            try:
                min_col, min_row, max_col, max_row = range_boundaries(rest)
                return self._expand_rect(
                    sheet, min_col, min_row, max_col, max_row
                )
            except Exception:
                return []

        # 单个单元格
        coord = _norm_coord(rest)
        if coord is not None:
            return [(sheet, coord)]

        # 整列 / 整列区域
        if _COL_RE.match(rest):
            parts = rest.split(":")
            c1 = _col_index(parts[0])
            c2 = _col_index(parts[-1])
            if c1 and c2:
                if c1 > c2:
                    c1, c2 = c2, c1
                _, max_row = self._sheet_dimensions(sheet)
                return self._expand_rect(sheet, c1, 1, c2, max_row)

        # 整行 / 整行区域
        if _ROW_RE.match(rest):
            parts = rest.split(":")
            r1 = _row_number(parts[0])
            r2 = _row_number(parts[-1])
            if r1 and r2:
                if r1 > r2:
                    r1, r2 = r2, r1
                max_col, _ = self._sheet_dimensions(sheet)
                return self._expand_rect(sheet, 1, r1, max_col, r2)

        # 结构化表引用 Table1[#All] / Table1[列名] / Table1[[列1],[列2]]
        if _STRUCTURED_RE.match(rest):
            return self._expand_structured(rest)

        # 定义名称（含 sheet 限定的）。名称不应包含 $（绝对标记），
        # 单元格 / 区域 / 行列引用在上面的分支已经处理。
        if "!" in raw or (
            rest
            and (rest[0].isalpha() or ord(rest[0]) > 127)
            and "$" not in rest
            and "." not in rest
        ):
            name = rest if "!" not in raw else f"{sheet}!{rest}"
            return self._resolve_defined_name(name, current_sheet, depth)

        return []

    def _expand_structured(self, rest: str) -> List[CellRef]:
        table_name = rest.split("[", 1)[0].strip()
        tbl = self._tables.get(table_name.casefold())
        if tbl is None:
            self._warn(f"未找到结构化表: {table_name}")
            return []

        sheet = self._table_sheets[table_name.casefold()]
        try:
            min_col, min_row, max_col, max_row = range_boundaries(tbl.ref)
        except Exception:
            return []

        spec = rest[rest.index("["):]  # 从第一个 [ 开始的整个说明符部分
        inner = spec[1:-1]  # 去掉最外层 [ ]
        # 说明符
        upper_spec = inner.upper()
        has_all = "#ALL" in upper_spec
        has_data_kw = any(k in upper_spec for k in ("#DATA", "#HEADERS", "#TOTALS"))
        # 收集列名：[单价] 或 [[单价],[数量]]
        col_names = [
            c.strip().strip("'")
            for c in re.findall(r"\[([^\[\]]+)\]", inner)
            if not c.strip().startswith("#")
        ]
        if not col_names and "[" not in inner and inner.strip():
            # 单括号形式：产品明细[单价]
            col_names = [inner.strip().strip("'")]

        header_row = getattr(tbl, "headerRowCount", 1) or 0
        data_start = min_row + (1 if header_row else 0)

        if has_all:
            return self._expand_rect(sheet, min_col, min_row, max_col, max_row)
        if has_data_kw:
            return self._expand_rect(sheet, min_col, data_start, max_col, max_row)
        if not col_names:
            # 裸表名 -> 数据行（不含表头）
            return self._expand_rect(sheet, min_col, data_start, max_col, max_row)

        names = {
            (getattr(c, "name", "") or "").casefold(): min_col + i
            for i, c in enumerate(getattr(tbl, "tableColumns", []))
        }
        out: List[CellRef] = []
        for cn in col_names:
            col = names.get(cn.casefold())
            if col is not None:
                out.extend(
                    self._expand_rect(sheet, col, data_start, col, max_row)
                )
        return out

    # -- 主流程 ------------------------------------------------------------ #
    def build(self) -> Analysis:
        self._collect_defined_names()

        for ws in self.wb.worksheets:
            self._scan_sheet(ws)

        # 节点按 (sheet, 行, 列) 稳定排序并重新编号
        ordered = sorted(
            self.nodes.values(),
            key=lambda n: (
                self._sheet_order(n.sheet),
                _coord_sort_key(n.coord),
            ),
        )
        nodes: List[Node] = []
        key_to_id: Dict[str, int] = {}
        for i, n in enumerate(ordered):
            n.id = i
            nodes.append(n)
            key_to_id[n.key] = i

        out: Dict[int, Set[int]] = {i: set() for i in range(len(nodes))}
        incoming: Dict[int, Set[int]] = {i: set() for i in range(len(nodes))}
        edges_out: Dict[str, List[str]] = {}
        for src_key, targets in self.edges.items():
            if src_key not in key_to_id:
                continue
            src_id = key_to_id[src_key]
            kept = []
            for t in sorted(targets):
                tid = key_to_id.get(t)
                if tid is None:
                    continue
                kept.append(t)
                out[src_id].add(tid)
                incoming[tid].add(src_id)
            if kept:
                edges_out[src_key] = kept

        cycles = find_cycles(out, nodes)

        formula_count = sum(1 for n in nodes if n.formula is not None)
        cross_sheet = sum(
            1
            for s, ts in edges_out.items()
            for t in ts
            if s.split("!", 1)[0] != t.split("!", 1)[0]
        )
        sheets_meta = [
            {
                "name": ws.title,
                "max_row": ws.max_row or 0,
                "max_column": ws.max_column or 0,
                "formula_cells": sum(
                    1
                    for n in nodes
                    if n.sheet == ws.title and n.formula is not None
                ),
            }
            for ws in self.wb.worksheets
        ]
        stats = {
            "sheet_count": len(sheets_meta),
            "node_count": len(nodes),
            "formula_count": formula_count,
            "edge_count": sum(len(v) for v in edges_out.values()),
            "cross_sheet_edge_count": cross_sheet,
            "external_ref_count": len(self.external_refs),
            "cycle_count": len(cycles),
            "defined_name_count": len(self.defined_names),
            "truncated_formula_count": self.truncated_formulas,
        }

        analysis = Analysis(
            filename=self.filename,
            sheets=sheets_meta,
            nodes=nodes,
            edges=edges_out,
            external_refs=sorted(self.external_refs),
            defined_names=self.defined_names,
            cycles=cycles,
            stats=stats,
            warnings=self.warnings,
        )
        analysis._key_to_id = key_to_id
        analysis._out = out
        analysis._in = incoming
        return analysis

    def _sheet_order(self, name: str) -> int:
        for i, ws in enumerate(self.wb.worksheets):
            if ws.title == name:
                return i
        return len(self.wb.worksheets)

    def _scan_sheet(self, ws: Worksheet) -> None:
        from openpyxl.cell.cell import Cell as XLCell

        for row in ws.iter_rows():
            for cell in row:
                if not isinstance(cell, XLCell) or cell.value is None:
                    continue
                coord = cell.coordinate  # openpyxl 已规范化为大写
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    formula = cell.value
                    node = self._node(ws.title, coord, formula=formula)
                    self._add_formula_edges(node, formula, ws.title, coord)
                else:
                    # 常量单元格：仅当它已经被引用过时才会存在；未被引用则跳过
                    existing = self.nodes.get((ws.title, coord))
                    if existing is not None and existing.value is None:
                        existing.value = _display_value(cell.value)
                    # 常量即使尚未存在也不预先建节点，避免把整张表都变成图节点；
                    # 引用它的公式会在展开时创建节点，这里在第二遍补齐 value。

        # 第二遍：为“被公式引用的常量单元格”补值
        for ref, node in self.nodes.items():
            sheet_name, c = ref
            if sheet_name != ws.title or node.value is not None or node.formula:
                continue
            xl = ws[c]
            if xl.value is not None and not (
                isinstance(xl.value, str) and xl.value.startswith("=")
            ):
                node.value = _display_value(xl.value)

    def _add_formula_edges(
        self, node: Node, formula: str, sheet: str, coord: str
    ) -> None:
        try:
            tokens = Tokenizer(formula).items
        except Exception:
            self._warn(f"无法解析公式 {sheet}!{coord}: {formula}")
            return

        targets: Set[CellRef] = set()
        truncated = False
        for tok in tokens:
            if tok.type != "OPERAND" or tok.subtype != "RANGE":
                continue
            expanded = self._expand_operand(tok.value, sheet)
            for ref in expanded:
                targets.add(ref)
                self._node(ref[0], ref[1])  # 确保被引用节点存在
                if len(targets) >= MAX_EDGES_PER_FORMULA:
                    truncated = True
                    break
            if truncated:
                break

        if truncated:
            self.truncated_formulas += 1
            self._warn(
                f"公式 {sheet}!{coord} 的引用数超过上限 "
                f"{MAX_EDGES_PER_FORMULA}，依赖已截断"
            )

        if self.edge_total + len(targets) > MAX_TOTAL_EDGES:
            self._warn(
                f"整个工作簿的依赖边数超过上限 {MAX_TOTAL_EDGES}，"
                f"公式 {sheet}!{coord} 的依赖未完整加入"
            )
            return

        key = node.key
        bucket = self.edges.setdefault(key, set())
        for ref in targets:
            tgt = self.nodes[ref].key
            if tgt not in bucket:
                bucket.add(tgt)
                self.edge_total += 1


def _display_value(v) -> str:
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    text = str(v)
    return text if len(text) <= 200 else text[:200] + "…"


def _coord_sort_key(coord: str) -> Tuple[int, int]:
    m = re.match(r"^([A-Z]+)([0-9]+)$", coord)
    if not m:
        return (1 << 30, 1 << 30)
    return (int(m.group(2)), column_index_from_string(m.group(1)))


# --------------------------------------------------------------------------- #
# 图算法
# ---------------------------------------------------------------------------
def find_cycles(adj: Dict[int, Set[int]], nodes: List[Node]) -> List[Cycle]:
    """迭代版 Tarjan 求强连通分量；非平凡 SCC 各提取一条代表环。"""
    index_at: Dict[int, int] = {}
    low: Dict[int, int] = {}
    on_stack: Set[int] = set()
    stack: List[int] = []
    sccs: List[List[int]] = []
    counter = 0

    for start in adj:
        if start in index_at:
            continue
        work: List[Tuple[int, Iterable[int]]] = [(start, iter(sorted(adj[start])))]
        index_at[start] = low[start] = counter
        counter += 1
        stack.append(start)
        on_stack.add(start)

        while work:
            v, it = work[-1]
            advanced = False
            for w in it:
                if w not in index_at:
                    index_at[w] = low[w] = counter
                    counter += 1
                    stack.append(w)
                    on_stack.add(w)
                    work.append((w, iter(sorted(adj[w]))))
                    advanced = True
                    break
                if w in on_stack:
                    low[v] = min(low[v], index_at[w])
            if advanced:
                continue

            work.pop()
            if work:
                parent = work[-1][0]
                low[parent] = min(low[parent], low[v])
            if low[v] == index_at[v]:
                comp: List[int] = []
                while True:
                    w = stack.pop()
                    on_stack.discard(w)
                    comp.append(w)
                    if w == v:
                        break
                if len(comp) > 1 or v in adj.get(v, ()):
                    sccs.append(comp)

    cycles: List[Cycle] = []
    for comp in sccs:
        members = set(comp)
        # 从最小节点出发，在 SCC 内部 DFS 找到一条回到起点的路径
        start = min(comp, key=lambda i: _coord_sort_key(nodes[i].coord))
        path = _cycle_path(start, members, adj)
        if path is None:
            continue
        keys = [nodes[i].key for i in path]
        keys.append(nodes[start].key)  # 闭环
        cycles.append(Cycle(nodes=keys))
    cycles.sort(key=lambda c: (len(c.nodes), c.nodes))
    return cycles


def _cycle_path(
    start: int, members: Set[int], adj: Dict[int, Set[int]]
) -> Optional[List[int]]:
    """在 members 诱导子图中找 start -> ... -> start 的简单环。"""
    if start in adj.get(start, ()):
        return [start]
    # DFS：栈元素为 (node, 邻居迭代器, 路径索引)
    path: List[int] = [start]
    on_path = {start}
    stack: List[Tuple[int, Iterable[int]]] = [
        (start, iter(sorted(n for n in adj[start] if n in members)))
    ]
    while stack:
        node, it = stack[-1]
        found_next = False
        for nb in it:
            if nb == start and len(path) > 1:
                return path
            if nb in on_path:
                continue
            path.append(nb)
            on_path.add(nb)
            stack.append(
                (nb, iter(sorted(n for n in adj[nb] if n in members)))
            )
            found_next = True
            break
        if not found_next:
            stack.pop()
            on_path.discard(path.pop())
    return None


def bfs_shortest(start: int, adj: Dict[int, Set[int]]) -> Dict[int, List[int]]:
    """BFS 返回 start 到各可达节点的最短路径（含起点）。"""
    paths: Dict[int, List[int]] = {start: [start]}
    q = deque([start])
    while q:
        v = q.popleft()
        for w in sorted(adj.get(v, ())):
            if w not in paths:
                paths[w] = paths[v] + [w]
                q.append(w)
    return paths


def enumerate_paths(
    start: int,
    goal: int,
    adj: Dict[int, Set[int]],
    max_paths: int = 3,
    max_depth: int = 20,
    budget: int = 20_000,
) -> List[List[int]]:
    """枚举 start -> goal 的简单路径（迭代 DFS），按长度升序。"""
    results: List[List[int]] = []
    path = [start]
    on_path = {start}
    steps = 0

    def dfs(v: int) -> None:
        nonlocal steps
        if len(results) >= max_paths or steps > budget:
            return
        if v == goal:
            results.append(list(path))
            return
        if len(path) - 1 >= max_depth:
            return
        for w in sorted(adj.get(v, ())):
            steps += 1
            if steps > budget:
                return
            if w in on_path:
                continue
            path.append(w)
            on_path.add(w)
            dfs(w)
            on_path.discard(w)
            path.pop()
            if len(results) >= max_paths:
                return

    dfs(start)
    results.sort(key=lambda p: (len(p), p))
    return results
