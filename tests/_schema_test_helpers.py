def _schema_contains_key(schema: object, key: str) -> bool:
    if isinstance(schema, dict):
        return key in schema or any(
            _schema_contains_key(value, key) for value in schema.values()
        )
    if isinstance(schema, list):
        return any(_schema_contains_key(value, key) for value in schema)
    return False


def _schema_contains_type_list(schema: object) -> bool:
    if isinstance(schema, dict):
        if isinstance(schema.get("type"), list):
            return True
        return any(_schema_contains_type_list(value) for value in schema.values())
    if isinstance(schema, list):
        return any(_schema_contains_type_list(value) for value in schema)
    return False
