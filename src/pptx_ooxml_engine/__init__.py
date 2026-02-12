"""pptx-ooxml-engine public API."""

from .engine import ApplyResult, apply_ops, generate_pptx
from .models import parse_ops, parse_plan
from .schema import load_ops_schema
from .verify import VerifyReport, verify_pptx

__version__ = "1.0.0"

__all__ = [
    "__version__",
    "ApplyResult",
    "VerifyReport",
    "apply_ops",
    "generate_pptx",
    "generate_example_outputs",
    "load_ops_schema",
    "parse_ops",
    "parse_plan",
    "verify_pptx",
]


def generate_example_outputs(*args, **kwargs):
    from .examples_runner import generate_example_outputs as _generate_example_outputs

    return _generate_example_outputs(*args, **kwargs)
