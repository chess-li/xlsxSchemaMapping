SCHEMA_SQL = """
create table if not exists standard_fields (
    key text primary key,
    display_name text not null,
    domain text not null,
    description text not null default ''
);

create table if not exists field_aliases (
    field_key text not null,
    alias text not null,
    unique(field_key, alias),
    foreign key(field_key) references standard_fields(key)
);

create table if not exists templates (
    signature text primary key,
    sheet_name text not null,
    mappings_json text not null,
    updated_at text not null default current_timestamp
);
"""


def initialize_schema(connection) -> None:
    connection.executescript(SCHEMA_SQL)
