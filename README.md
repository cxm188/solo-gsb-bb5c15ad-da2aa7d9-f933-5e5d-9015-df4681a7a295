# 电子表格公式影响分析（Formula Impact Analyzer）

一个可自托管的 Web 应用：上传 Excel（`.xlsx` / `.xlsm`），自动解析公式中的
**跨工作表引用**与**矩形区域引用**，展开为单元格级依赖图，检测**循环引用**，
并对任意单元格列出所有**直接 / 间接受影响单元格及传播路径**，支持按工作表筛选、
查看公式与依赖链、导出完整分析 JSON。

- 后端：FastAPI + openpyxl（不计算公式，只做静态引用分析）
- 前端：原生 HTML / CSS / JavaScript，无构建步骤
- 部署：Dockerfile / Docker Compose 一键启动
- 数据只在服务进程内存中处理，上传文件不会写入磁盘

## 功能一览

| 功能 | 说明 |
| --- | --- |
| 跨表引用 | `Sheet2!A1`、`'My Sheet'!$B$2` 均能解析，工作表名不区分大小写 |
| 区域引用 | `A1:B10` 展开为区域内每个单元格；支持整列 `B:B`、整行 `3:5`（按工作表维度截断） |
| 定义名称 | 工作簿级 / 工作表级名称，名称指向区域或公式（递归 3 层） |
| 结构化表 | `Table1[列名]`、`Table1[[列1],[列2]]`、`Table1[#All]` / `[#Data]` |
| 外部引用 | `[1]Sheet1!A1` 等跨工作簿引用会被单独标记，不参与图展开 |
| 循环检测 | Tarjan 强连通分量算法，给出自环与多单元格环的完整回路 |
| 影响分析 | 两个方向：**下游影响**（谁会因它变化而变化）/ **上游依赖**（它依赖谁） |
| 路径展示 | 每个受影响格给出最短传播链；图较小时额外列出多条不同路径 |
| 筛选与检索 | 按工作表过滤、仅看公式格、按坐标 / 公式文本搜索 |
| JSON 导出 | 导出包含节点、边、循环、统计、警告的完整分析结果 |

## 快速开始（Docker Compose）

```bash
docker compose up --build
```

启动后浏览器打开：

```
http://localhost:8000
```

健康检查：`GET http://localhost:8000/api/workbooks`。

停止：

```bash
docker compose down
```

## 本地直接运行（不用 Docker）

需要 Python 3.11+：

```bash
python3 -m pip install -r requirements.txt
python3 -m uvicorn app.main:app --host 0.0.0.0 --port 8000
# 打开 http://localhost:8000
```

## 端口与环境变量

| 变量 | 默认值 | 说明 |
| --- | --- | --- |
| `8000`（容器端口） | — | Compose 映射为宿主机 `8000:8000`，改端口编辑 `docker-compose.yml` 左侧端口 |
| `MAX_UPLOAD_MB` | `20` | 上传文件大小上限（MB） |
| `MAX_AFFECTED` | `20000` | 单次影响分析最多返回的受影响单元格数量 |

## 示例文件

仓库自带一个覆盖全部场景的示例工作簿：

```
examples/sample_formulas.xlsx
```

重新生成（修改后）：

```bash
python3 examples/make_sample.py
```

示例包含 5 个工作表：

1. **基础数据** — 月度销售/成本，`SUM(矩形区域)`、`AVERAGE(B:B)` 整列引用；
2. **季度汇总** — 跨表引用基础数据，并使用定义名称 `税率`；
3. **管理层报告** — 引用汇总表，与前两表形成 6 跳的多级影响链
   （修改 `基础数据!B2` 可一直影响到 `管理层报告!B4`）；
4. **产品表** — 结构化表 `产品明细[单价]` / `[数量]` 列级引用；
5. **循环引用演示** — 一个三单元格环 `A3→B3→C3→A3` 与一个自环 `B6→B6`。

页面上传该文件后，可在「单元格与公式」中点击 `基础数据!B2` 的 **分析影响**，
观察完整的跨表传播路径与循环标记。

## 使用说明

1. 打开首页，拖拽或选择 `.xlsx` 文件上传；
2. 顶部统计区显示工作表数、公式数、依赖边、跨表引用、循环数等；
3. 左侧面板：按工作表筛选 / 搜索 / 仅看公式，点击 **分析影响**；
4. 右侧面板切换「下游影响 / 上游依赖」：
   - **直接**：一跳（直接引用 / 直接被引用）；
   - **间接**：多跳，按距离排序，每条给出类似
     `基础数据!B2 → 基础数据!F1 → 季度汇总!B2 → 管理层报告!B2` 的传播链；
   - 处在循环上的单元格会有「循环」标记；
5. 底部「检测到的循环引用」列出所有回路（首尾相同表示闭环）；
6. 右上角 **导出分析 JSON** 下载完整结果。

## 导出 JSON 结构

`GET /api/workbooks/{id}/export`（或页面按钮）：

```json
{
  "filename": "sample_formulas.xlsx",
  "sheets": [{ "name": "基础数据", "max_row": 7, "max_column": 6, "formula_cells": 5 }],
  "nodes": [{ "id": 0, "key": "基础数据!B2", "sheet": "基础数据", "coord": "B2",
              "value": "1200", "formula": null }],
  "edges": { "基础数据!F1": ["基础数据!B2", "基础数据!B3"] },
  "external_refs": [],
  "defined_names": { "税率": "0.13" },
  "cycles": [["循环引用演示!A3", "循环引用演示!B3", "循环引用演示!C3", "循环引用演示!A3"]],
  "stats": { "node_count": 38, "edge_count": 44, "cycle_count": 2 },
  "warnings": []
}
```

## HTTP API

| 方法 | 路径 | 说明 |
| --- | --- | --- |
| `POST` | `/api/workbooks` | 上传 XLSX（multipart 字段 `file`），返回 id 与整体分析摘要 |
| `GET` | `/api/workbooks/{id}` | 获取完整分析 |
| `GET` | `/api/workbooks/{id}/nodes?sheet=&formula_only=&q=` | 单元格列表（可筛选 / 搜索） |
| `GET` | `/api/workbooks/{id}/impact?sheet=&coord=&direction=` | 影响分析，`direction=dependents`（默认）或 `precedents` |
| `GET` | `/api/workbooks/{id}/export` | 导出分析 JSON（附件下载） |
| `DELETE` | `/api/workbooks/{id}` | 从内存中移除该工作簿 |
| `GET` | `/api/workbooks` | 已上传工作簿列表（健康检查用） |

影响分析响应中的 `affected[].paths` 为路径数组（元素为含 `key/sheet/coord/formula`
的节点，路径第 1 个是所选单元格、最后一个是目标），`distance` 为传播跳数。

## 测试

```bash
python3 examples/make_sample.py     # 先确保示例文件存在
python3 -m unittest discover -s tests -v
```

## 实现说明与边界

- 只做**静态引用解析**（基于 openpyxl 的公式 tokenizer），不执行公式，因此没有宏 / 注入风险；
- 单个区域展开上限 100,000 单元格、单公式边上限 200,000、整图边上限 2,000,000，
  超出会截断并在页面「警告」中提示（防止整行 / 整列引用造成内存膨胀）；
- 上传的工作簿保存在进程内存（按随机 id 索引），重启服务即清空；
- 暂不支持数组公式的隐式溢出区域、PivotTable 与 INDIRECT 等动态引用（后者只有运行时才能确定目标）。
