from __future__ import annotations

import subprocess
import sys
from pathlib import Path

from pptx import Presentation


def test_generate_example_outputs_creates_two_valid_pptx(tmp_path: Path) -> None:
    from pptx_ooxml_engine.examples_runner import generate_example_outputs

    outputs = generate_example_outputs(tmp_path)
    assert len(outputs) == 2
    for output in outputs:
        assert output.exists()
        prs = Presentation(str(output))
        assert len(prs.slides) >= 1


def test_examples_runner_module_cli(tmp_path: Path) -> None:
    project_root = Path(__file__).resolve().parents[1]
    completed = subprocess.run(
        [
            sys.executable,
            "-m",
            "pptx_ooxml_engine.examples_runner",
            "--output-dir",
            str(tmp_path),
        ],
        cwd=str(project_root),
        env={"PYTHONPATH": str(project_root / "src")},
        capture_output=True,
        text=True,
        check=False,
    )
    assert completed.returncode == 0, completed.stderr
    assert (tmp_path / "example_01.pptx").exists()
    assert (tmp_path / "example_02.pptx").exists()
