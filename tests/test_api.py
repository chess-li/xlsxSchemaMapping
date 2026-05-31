from io import BytesIO

from openpyxl import Workbook
from starlette.testclient import TestClient

from excel_matcher.api import app as api_module
from excel_matcher.storage.sqlite_store import SQLiteStore


def _workbook_bytes() -> bytes:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "订单"
    sheet["A1"] = "购买方"
    sheet["B1"] = "订单金额"
    sheet["A2"] = "腾讯"
    sheet["B2"] = 1000
    buffer = BytesIO()
    workbook.save(buffer)
    return buffer.getvalue()


def test_health_reports_components():
    client = TestClient(api_module.app)

    response = client.get("/health")

    assert response.status_code == 200
    body = response.json()
    assert set(body["components"]) == {"sqlite", "embedding", "llm"}
    assert body["components"]["sqlite"]["available"] is True
    assert "reason" in body["components"]["embedding"]
    assert "reason" in body["components"]["llm"]


def test_fields_endpoint_returns_built_in_dictionary():
    client = TestClient(api_module.app)

    response = client.get("/fields")

    assert response.status_code == 200
    fields = response.json()["fields"]
    customer = next(field for field in fields if field["key"] == "customer_name")
    assert customer["display_name"] == "客户名称"
    assert "购买方" in customer["aliases"]


def test_workbooks_analyze_accepts_xlsx_upload():
    client = TestClient(api_module.app)

    response = client.post(
        "/workbooks/analyze",
        files={
            "file": (
                "order.xlsx",
                _workbook_bytes(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        },
    )

    assert response.status_code == 200
    body = response.json()
    assert body["workbook"]["sheets"][0]["name"] == "订单"
    assert body["headers"]["订单"]["header_row"] == 1
    assert [profile["column_name"] for profile in body["profiles"]["订单"]] == [
        "购买方",
        "订单金额",
    ]


def test_fields_match_endpoint_returns_mapping_results():
    client = TestClient(api_module.app)

    response = client.post(
        "/fields/match",
        json={
            "profiles": [
                {
                    "column_name": "购买方",
                    "column_index": 1,
                    "data_type": "TEXT",
                    "null_rate": 0.0,
                    "unique_rate": 1.0,
                    "samples": ["腾讯"],
                }
            ],
            "enable_embedding": False,
            "enable_llm": False,
        },
    )

    assert response.status_code == 200
    result = response.json()["mappings"][0]
    assert result["target_field"] == "customer_name"
    assert [stage["stage"] for stage in result["stage_results"]] == [
        "exact",
        "dictionary",
        "rapidfuzz",
        "embedding",
        "llm",
    ]


def test_templates_endpoint_saves_confirmed_mappings(tmp_path):
    api_module.app.state.sqlite_path = tmp_path / "templates.db"
    client = TestClient(api_module.app)

    response = client.post(
        "/templates",
        json={
            "signature": "same-structure",
            "sheet_name": "订单",
            "mappings": {"购买方": "customer_name"},
        },
    )

    assert response.status_code == 200
    assert response.json() == {"saved": True}
    assert SQLiteStore(api_module.app.state.sqlite_path).find_template("same-structure")[
        "mappings"
    ] == {"购买方": "customer_name"}
