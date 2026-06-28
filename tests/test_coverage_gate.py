"""Tests for FR-17 coverage gate and CI configuration (issue #44).

AC2: pyproject.toml에 --cov-fail-under=75 설정 존재 단언
AC4: core-only 환경에서 pytest 실패 없음 (NFR-08)
AC6: CI 워크플로에 커버리지 게이트 포함 확인

AC-to-test mapping:
  AC2 -> test_ac2_cov_fail_under_in_pyproject
  AC4 -> test_ac4_core_only_no_vlm_sdk_needed
  AC6 -> test_ac6_ci_yml_runs_pytest_with_cov_gate
"""

from __future__ import annotations

import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).parent.parent


class TestAc2CovFailUnderConfig:
    """ac2_pyproject_cov_fail_under_설정 — pyproject.toml must have --cov-fail-under."""

    def test_ac2_cov_fail_under_in_pyproject(self) -> None:
        """AC2: pyproject.toml addopts must contain --cov-fail-under=75."""
        pyproject_path = PROJECT_ROOT / "pyproject.toml"
        assert pyproject_path.exists(), "pyproject.toml must exist"

        content = pyproject_path.read_text(encoding="utf-8")
        assert (
            "--cov-fail-under=75" in content
        ), "pyproject.toml must contain '--cov-fail-under=75' in addopts (ADR-605)"

    def test_ac2_cov_fail_under_in_addopts_section(self) -> None:
        """AC2: --cov-fail-under=75 must be in [tool.pytest.ini_options] addopts."""
        try:
            import tomllib  # Python 3.11+ stdlib
        except ImportError:
            import tomli as tomllib  # type: ignore[no-redef]

        pyproject_path = PROJECT_ROOT / "pyproject.toml"
        with pyproject_path.open("rb") as f:
            data = tomllib.load(f)

        addopts = (
            data.get("tool", {})
            .get("pytest", {})
            .get("ini_options", {})
            .get("addopts", "")
        )
        assert "--cov-fail-under=75" in addopts, (
            f"[tool.pytest.ini_options].addopts must contain --cov-fail-under=75, "
            f"got: {addopts!r}"
        )

    def test_ac2_cov_source_also_set(self) -> None:
        """AC2: addopts must also include --cov=src/pptx_md for the gate to apply."""
        try:
            import tomllib
        except ImportError:
            import tomli as tomllib  # type: ignore[no-redef]

        pyproject_path = PROJECT_ROOT / "pyproject.toml"
        with pyproject_path.open("rb") as f:
            data = tomllib.load(f)

        addopts = (
            data.get("tool", {})
            .get("pytest", {})
            .get("ini_options", {})
            .get("addopts", "")
        )
        assert (
            "--cov=src/pptx_md" in addopts
        ), f"addopts must include '--cov=src/pptx_md', got: {addopts!r}"


class TestAc4CoreOnlyNFR08:
    """ac4_core_only_no_fail — pytest must pass without VLM SDKs (NFR-08)."""

    def test_ac4_convert_importable_without_vlm_sdk(self) -> None:
        """AC4: convert() is importable even when anthropic/openai SDKs are absent."""
        # Simulate SDK absence: inject None (triggers ImportError on import)
        prev_anthropic = sys.modules.get("anthropic", "ABSENT")
        prev_openai = sys.modules.get("openai", "ABSENT")

        sys.modules["anthropic"] = None  # type: ignore[assignment]
        sys.modules["openai"] = None  # type: ignore[assignment]

        try:
            # Reload pptx_md.api to verify it does not import VLM SDKs
            import importlib

            import pptx_md.api as api_mod

            importlib.reload(api_mod)
            # If we reach here, api.py does not eagerly import VLM SDKs
            assert hasattr(api_mod, "convert"), "convert must exist in api module"
        finally:
            if prev_anthropic == "ABSENT":
                sys.modules.pop("anthropic", None)
            else:
                sys.modules["anthropic"] = prev_anthropic  # type: ignore[assignment]
            if prev_openai == "ABSENT":
                sys.modules.pop("openai", None)
            else:
                sys.modules["openai"] = prev_openai  # type: ignore[assignment]

    def test_ac4_no_vlm_sdk_in_sys_modules_after_import(self) -> None:
        """AC4 (NFR-08): importing pptx_md must NOT bring in VLM SDKs."""
        # This test verifies the already-imported state (VLM SDKs absent = good)
        import pptx_md  # noqa: F401

        assert (
            "anthropic" not in sys.modules
        ), "anthropic must not be in sys.modules after `import pptx_md`"
        assert (
            "openai" not in sys.modules
        ), "openai must not be in sys.modules after `import pptx_md`"


class TestAc6CiGate:
    """ac6_ci_yml_커버리지_게이트 — CI workflow must run pytest (applies gate)."""

    def test_ac6_ci_yml_exists(self) -> None:
        """AC6: .github/workflows/ci.yml must exist."""
        ci_yml = PROJECT_ROOT / ".github" / "workflows" / "ci.yml"
        assert ci_yml.exists(), ".github/workflows/ci.yml must exist"

    def test_ac6_ci_yml_runs_pytest(self) -> None:
        """AC6: ci.yml must contain a pytest step that triggers the coverage gate."""
        ci_yml = PROJECT_ROOT / ".github" / "workflows" / "ci.yml"
        content = ci_yml.read_text(encoding="utf-8")
        assert "pytest" in content, "ci.yml must contain a pytest step"

    def test_ac6_ci_gate_applies_via_pyproject(self) -> None:
        """AC6: the coverage gate is applied because pyproject.toml addopts is set
        and CI runs bare 'pytest' which picks up addopts automatically."""
        ci_yml = PROJECT_ROOT / ".github" / "workflows" / "ci.yml"
        content = ci_yml.read_text(encoding="utf-8")

        # CI must run pytest (gate comes from pyproject.toml addopts — ADR-605)
        assert "pytest" in content, "CI must run pytest"

        # pyproject.toml must have the gate (already verified in AC2 tests,
        # but cross-check here as the AC6 evidence anchor)
        pyproject_path = PROJECT_ROOT / "pyproject.toml"
        pyproject_content = pyproject_path.read_text(encoding="utf-8")
        assert (
            "--cov-fail-under=75" in pyproject_content
        ), "Coverage gate must be in pyproject.toml so that CI pytest picks it up"
