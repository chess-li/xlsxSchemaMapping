import json
import sqlite3
from pathlib import Path
from typing import Any

from excel_matcher.models import StandardField
from excel_matcher.storage.default_fields import get_default_fields
from excel_matcher.storage.schema import initialize_schema


class SQLiteStore:
    def __init__(self, path: str | Path):
        self.path = Path(path)

    def initialize(self) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        with self._connect() as connection:
            initialize_schema(connection)

    def seed_default_fields(self) -> None:
        with self._connect() as connection:
            initialize_schema(connection)
            for field in get_default_fields():
                connection.execute(
                    """
                    insert into standard_fields(key, display_name, domain, description)
                    values (?, ?, ?, ?)
                    on conflict(key) do update set
                        display_name = excluded.display_name,
                        domain = excluded.domain,
                        description = excluded.description
                    """,
                    (field.key, field.display_name, field.domain, field.description),
                )
                for alias in field.aliases:
                    connection.execute(
                        """
                        insert or ignore into field_aliases(field_key, alias)
                        values (?, ?)
                        """,
                        (field.key, alias),
                    )

    def list_standard_fields(self) -> list[StandardField]:
        with self._connect() as connection:
            rows = connection.execute(
                """
                select key, display_name, domain, description
                from standard_fields
                order by key
                """
            ).fetchall()
            alias_rows = connection.execute(
                """
                select field_key, alias
                from field_aliases
                order by rowid
                """
            ).fetchall()

        aliases_by_key: dict[str, list[str]] = {}
        for field_key, alias in alias_rows:
            aliases_by_key.setdefault(field_key, []).append(alias)
        return [
            StandardField(
                key=row["key"],
                display_name=row["display_name"],
                domain=row["domain"],
                description=row["description"],
                aliases=aliases_by_key.get(row["key"], []),
            )
            for row in rows
        ]

    def save_template(
        self,
        signature: str,
        sheet_name: str,
        mappings: dict[str, str],
    ) -> None:
        with self._connect() as connection:
            initialize_schema(connection)
            connection.execute(
                """
                insert into templates(signature, sheet_name, mappings_json, updated_at)
                values (?, ?, ?, current_timestamp)
                on conflict(signature) do update set
                    sheet_name = excluded.sheet_name,
                    mappings_json = excluded.mappings_json,
                    updated_at = current_timestamp
                """,
                (
                    signature,
                    sheet_name,
                    json.dumps(mappings, ensure_ascii=False, sort_keys=True),
                ),
            )

    def find_template(self, signature: str) -> dict[str, Any] | None:
        with self._connect() as connection:
            initialize_schema(connection)
            row = connection.execute(
                """
                select signature, sheet_name, mappings_json
                from templates
                where signature = ?
                """,
                (signature,),
            ).fetchone()
        if row is None:
            return None
        return {
            "signature": row["signature"],
            "sheet_name": row["sheet_name"],
            "mappings": json.loads(row["mappings_json"]),
        }

    def _connect(self):
        connection = sqlite3.connect(self.path)
        connection.row_factory = sqlite3.Row
        return connection
