"""HTTP API 端到端测试：上传 -> 查询影响面 -> 导出。"""
import io

from fastapi.testclient import TestClient
from openpyxl import Workbook

from app.main import app

client = TestClient(app)


def make_xlsx() -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "S1"
    ws["A1"] = 3
    ws["B1"] = "=A1*2"
    ws2 = wb.create_sheet("S2")
    ws2["C3"] = "=S1!B1+1"
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def test_health():
    assert client.get("/api/health").json()["status"] == "ok"


def test_index_page():
    r = client.get("/")
    assert r.status_code == 200 and "影响分析" in r.text


def test_upload_and_impact_flow():
    r = client.post(
        "/api/upload",
        files={"file": ("t.xlsx", make_xlsx(),
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )
    assert r.status_code == 200
    d = r.json()
    wid = d["workbook_id"]
    assert d["formula_count"] == 2
    assert any(s["name"] == "S1" for s in d["sheets"])

    r = client.get(f"/api/workbooks/{wid}/impact",
                   params={"cell": "S1!A1", "direction": "downstream"})
    assert r.status_code == 200
    imp = r.json()
    cells = {i["cell"] for i in imp["impacted"]}
    assert {"S1!B1", "S2!C3"} <= cells

    # 上游反向查询
    r = client.get(f"/api/workbooks/{wid}/impact",
                   params={"cell": "S2!C3", "direction": "upstream"})
    cells = {i["cell"] for i in r.json()["impacted"]}
    assert {"S1!B1", "S1!A1"} <= cells

    # 工作表筛选
    r = client.get(f"/api/workbooks/{wid}/cells", params={"sheet": "S2"})
    assert r.status_code == 200
    assert all(c["sheet"] == "S2" for c in r.json()["cells"])

    # 导出
    r = client.get(f"/api/workbooks/{wid}/export")
    assert r.status_code == 200
    export = r.json()
    assert export["filename"] == "t.xlsx"
    assert len(export["edges"]) == 2
    assert {"from": "S1!A1", "to": "S1!B1"} in export["edges"]
    assert "attachment" in r.headers["content-disposition"]


def test_bad_upload():
    r = client.post("/api/upload", files={"file": ("t.txt", b"hello", "text/plain")})
    assert r.status_code == 400


def test_missing_workbook():
    r = client.get("/api/workbooks/nonexistent")
    assert r.status_code == 404
