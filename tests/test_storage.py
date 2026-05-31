import sqlite3

from excel_matcher.storage.sqlite_store import SQLiteStore


def test_sqlite_store_initializes_and_seeds_default_fields_idempotently(tmp_path):
    db_path = tmp_path / "excel_matcher.db"
    store = SQLiteStore(db_path)

    store.initialize()
    store.seed_default_fields()
    store.seed_default_fields()

    with sqlite3.connect(db_path) as connection:
        tables = {
            row[0]
            for row in connection.execute(
                "select name from sqlite_master where type = 'table'"
            )
        }
        alias_count = connection.execute(
            "select count(*) from field_aliases where alias = '购买方'"
        ).fetchone()[0]

    fields = store.list_standard_fields()
    customer = next(field for field in fields if field.key == "customer_name")

    assert {"standard_fields", "field_aliases", "templates"} <= tables
    assert alias_count == 1
    assert customer.display_name == "客户名称"
    assert "购买方" in customer.aliases
    assert "企业名称" in customer.aliases


def test_sqlite_store_upserts_and_finds_templates(tmp_path):
    store = SQLiteStore(tmp_path / "excel_matcher.db")
    store.initialize()

    store.save_template(
        signature="same-structure",
        sheet_name="订单",
        mappings={"购买方": "customer_name"},
    )
    store.save_template(
        signature="same-structure",
        sheet_name="订单更新",
        mappings={"购买方": "customer_name", "金额": "amount"},
    )

    template = store.find_template("same-structure")

    assert template == {
        "signature": "same-structure",
        "sheet_name": "订单更新",
        "mappings": {"购买方": "customer_name", "金额": "amount"},
    }
    assert store.find_template("missing") is None
