from __future__ import annotations

import json
from importlib.resources import files


def load_ops_schema(version: str = "v1") -> dict:
    schema_path = files("pptx_ooxml_engine").joinpath(f"schemas/ops.{version}.json")
    with schema_path.open("r", encoding="utf-8") as f:
        return json.load(f)
