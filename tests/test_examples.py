from __future__ import annotations

import json
from pathlib import Path


def test_example_ops_files_are_parseable() -> None:
    from pptx_ooxml_engine.models import parse_plan

    project_root = Path(__file__).resolve().parents[1]
    example_dir = project_root / "examples" / "ops"
    example_files = sorted(example_dir.glob("*.json"))

    assert example_files, "no example ops json files found"
    for example_file in example_files:
        data = json.loads(example_file.read_text(encoding="utf-8"))
        plan = parse_plan(data)
        assert isinstance(plan.operations, list)
        assert len(plan.operations) > 0
