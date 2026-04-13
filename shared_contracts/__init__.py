from __future__ import annotations

import json
from functools import lru_cache
from importlib.resources import files
from typing import Any


@lru_cache(maxsize=None)
def load_json_contract(name: str) -> dict[str, Any]:
    return json.loads(files(__package__).joinpath(name).read_text(encoding="utf-8"))
