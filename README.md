# 电子表格公式影响分析（Spreadsheet Formula Impact Analysis）

一个可自托管的 Web 应用：上传 Excel（`.xlsx`）文件后，系统解析工作表中的公式，
将**跨工作表引用**与**矩形区域引用**展开为单元格级别的依赖图，自动检测**循环引用**，
并可针对任意单元格列出其**所有直接 / 间接受影响单元格及完整影响路径**。

- 后端：FastAPI + openpyxl
- 前端：原生 HTML / CSS / JavaScript 单页（无需构建，无需 Node）
- 部署：Dockerfile + docker-compose 一键启动

## 功能一览

- 📤 拖拽 / 点选上传 `.xlsx`（上限 25MB，内存中保留最近 5 个文件）
- 🔗 解析单元格引用、矩形区域（`A1:B10`）、整列（`B:B`）、整行（`2:5`）、
  绝对引用（`$A$1`）、跨表引用（`参数!B2`）、带空格的工作表名（`'月报 2026'!B2`）
- 🔁 Tarjan 强连通分量算法检测循环引用（含自环），页面顶部高亮展示
- 🔍 选定单元格后展示：
  - **下游影响**：哪些单元格直接 / 间接引用了它（改了它会影响谁）
  - **上游依赖**：该单元格直接 / 间接依赖谁（它的计算需要谁）
  - 每个受影响单元格的**层级（深度）**和一条**最短影响路径**
- 🗂 可按工作表筛选、按单元格 / 公式文本搜索、切换是否只看公式单元格
- 📄 查看每个单元格的公式原文与 Excel 缓存计算值
- 💾 一键导出完整分析 JSON（工作表统计、公式清单、依赖边、循环环）

## 一键启动（Docker Compose）

```bash
docker compose up --build -d
```

然后浏览器打开：

> **http://localhost:8000**

停止：

```bash
docker compose down
```

如需修改端口，编辑 `docker-compose.yml` 中的 `"8000:8000"`（左为主机端口）。

## 本地开发（不使用 Docker）

需要 Python 3.11+：

```bash
pip install -r requirements.txt
python scripts/make_sample.py          # 生成示例文件（可选，仓库已附带）
uvicorn app.main:app --host 0.0.0.0 --port 8000 --reload
```

运行测试：

```bash
pip install pytest httpx
python -m pytest tests/ -v
```

## 端口与 API

服务监听 **8000** 端口，全部接口：

| 方法 | 路径 | 说明 |
|---|---|---|
| `GET` | `/` | 前端单页 |
| `POST` | `/api/upload` | 上传 xlsx（multipart 字段名 `file`），返回工作簿 id、统计与循环信息 |
| `GET` | `/api/workbooks/{id}` | 工作簿统计摘要 |
| `GET` | `/api/workbooks/{id}/cells?sheet=&q=&formulas_only=` | 单元格列表（支持工作表筛选 / 搜索） |
| `GET` | `/api/workbooks/{id}/impact?cell=参数!B2&direction=downstream` | 影响面分析；`direction` 取 `downstream`（默认，受影响）或 `upstream`（依赖链） |
| `GET` | `/api/workbooks/{id}/export` | 导出完整分析 JSON |
| `GET` | `/samples/sample_impact.xlsx` | 下载示例文件 |
| `GET` | `/api/health` | 健康检查 |

### 导出 JSON 结构

```json
{
  "filename": "sample_impact.xlsx",
  "summary": { "sheets": [...], "formula_count": 18, "edge_count": 43,
               "cycle_count": 2, "cycles": [["循环!A1","循环!B1"],["循环!C1"]] },
  "formulas": { "计算!B2": { "formula": "=B2*参数!B4", "value": null, "in_cycle": false } },
  "edges": [ { "from": "参数!B2", "to": "计算!B2" } ],
  "cycles": [ ["循环!A1", "循环!B1"], ["循环!C1"] ]
}
```

> 说明：`value` 取自 Excel 保存时的缓存计算结果。openpyxl 不执行公式计算，
> 由程序新建、从未在 Excel 中打开保存的文件其缓存值为 `null`，属正常现象。

## 示例文件

`samples/sample_impact.xlsx`（由 `scripts/make_sample.py` 生成，也可在首页直接下载）
包含 5 个工作表，覆盖所有分析场景：

| 工作表 | 内容 |
|---|---|
| `参数` | 纯输入项（单价、数量、折扣率、税率、月数） |
| `计算` | 跨表引用链：单价×数量 → 折扣 → 税额 → 月合计 → 年合计 |
| `汇总` | 跨表矩形区域 `SUM(计算!B2:B5)`、整列引用 `SUM(参数!B:B)` |
| `月报 2026` | 带空格的工作表名（公式里写作 `='月报 2026'!...`） |
| `循环` | 故意制造的循环引用：`A1↔B1` 互引、`C1=C1*2` 自环、`D1` 为环的下游 |

推荐体验路径：

1. 首页下载并上传 `sample_impact.xlsx`，顶部可看到红色循环引用提示；
2. 左侧选择 `参数!B2`（单价），右侧显示 14 个受影响单元格与逐层路径；
3. 切换到「上游依赖」，选择 `'月报 2026'!B3`，可看到完整回溯链直到 `参数!B2`；
4. 用工作表筛选定位 `循环!A1`，观察循环标记与影响面的截断行为；
5. 点击「导出分析 JSON」获得机器可读的完整结果。

## 实现要点

- `app/formula_parser.py`：词法式正则扫描公式，区分字符串字面量 / `#REF!` 错误 /
  函数名，避免把 `LOG10`、`"A1"` 误判为单元格；输出未展开的 `RawRef`。
- `app/graph.py`：
  - 加载两次工作簿（`data_only=False` 取公式、`data_only=True` 取缓存值）；
  - 区域展开为具体单元格，整行 / 整列按工作表已用范围截断，单个区域上限 5000 格
    （`MAX_RANGE_CELLS`，防 `SUM(A:XFD)` 之类打爆内存）；
  - 迭代版 Tarjan SCC 检测循环引用；
  - BFS 求影响面，记录最短路径、层级与环标记，默认最大深度 50 / 5000 节点。
- `app/main.py`：FastAPI 路由与内存 LRU 式工作簿缓存（不落盘，重启即清空）。

## 目录结构

```
.
├── app/
│   ├── main.py              # FastAPI 入口与 API
│   ├── formula_parser.py    # 公式引用解析
│   ├── graph.py             # 依赖图 / 环检测 / 影响面 BFS
│   └── static/index.html    # 前端单页
├── samples/sample_impact.xlsx
├── scripts/make_sample.py   # 示例文件生成脚本
├── tests/                   # pytest 用例（22 个）
├── Dockerfile
├── docker-compose.yml
└── requirements.txt
```
