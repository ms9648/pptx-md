"""M1 smoke test — guarantees the package imports and the CI pipeline runs.

Covers FR-01 AC2 (import succeeds) and gives pytest a target so the CI
``pytest`` step exercises a real (passing) test.
"""


def test_ac2_import_pptx_md() -> None:
    """AC2: import pptx_md must succeed without ImportError."""
    import pptx_md

    assert pptx_md is not None
