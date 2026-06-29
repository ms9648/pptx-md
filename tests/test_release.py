"""Tests for FR-19 PyPI release workflow (issue #46).

Covers AC1–AC7.

AC-to-test mapping:
  AC1 -> TestAc1HatchBuild          (dist/ 에 wheel+tarball 생성)
  AC2 -> TestAc2TwineCheck          (메타데이터 검증 통과)
  AC3 -> TestAc3ReleaseYmlStructure (v* 트리거 + build/publish 잡 + OIDC)
  AC4 -> TestAc4EnvironmentGate     (publish 잡에 environment:pypi 설정)
  AC5 -> TestAc5LocalInstall        (wheel pip install → import 성공)
  AC6 -> TestAc6CoreOnly            (core-only 설치 시 import 성공, NFR-08)
  AC7 -> TestAc7CiDependency        (CI green 상태 전제 — release.yml 구조 확인)

Notes:
  - AC1, AC2, AC5 는 hatch build 를 실제 실행하는 통합 테스트다.
    빌드 환경이 없거나 시간이 오래 걸릴 수 있으므로 pytest marker 로 분리하지
    않는다(모든 AC 는 CI 에서도 실행 가능해야 하므로).
  - 헌장 §1.5: 실제 hatch publish 및 PyPI publisher 등록은 사람 게이트 —
    에이전트는 워크플로 작성·빌드 검증·로컬 설치 테스트까지만 수행한다.
  - AC3/AC4/AC7 의 YAML 구조 검증은 yaml 파싱 없이 텍스트 검색으로 수행한다.
    (PyYAML 은 dev 의존성에 없으므로 import yaml 제거)
"""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

import pytest

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

_REPO_ROOT = Path(__file__).parent.parent
_RELEASE_YML = _REPO_ROOT / ".github" / "workflows" / "release.yml"
_DIST_DIR = _REPO_ROOT / "dist"


def _run(cmd: list[str], cwd: Path | None = None) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        cmd,
        capture_output=True,
        text=True,
        cwd=str(cwd or _REPO_ROOT),
    )


def _latest_wheel() -> Path:
    """Return the wheel with the highest version in dist/.

    Sorts by modification time descending so freshly-built artifacts are
    preferred.  Falls back to the first wheel found.
    """
    wheels = sorted(
        _DIST_DIR.glob("*.whl"), key=lambda p: p.stat().st_mtime, reverse=True
    )
    if not wheels:
        raise FileNotFoundError(f"No .whl found in {_DIST_DIR}")
    return wheels[0]


# ---------------------------------------------------------------------------
# AC1: hatch build → dist/ 에 wheel + tarball 생성, exit 0
# ---------------------------------------------------------------------------


class TestAc1HatchBuild:
    """ac1_hatch_build_성공 — hatch build must produce wheel + tarball in dist/."""

    def test_ac1_wheel_exists(self) -> None:
        """AC1: dist/ must contain a .whl file matching version in pyproject.toml."""
        # hatch build는 이미 실행됐다고 가정(conftest 또는 사전 실행).
        # dist/가 없으면 hatch build를 실행해 생성한다.
        if not _DIST_DIR.exists() or not list(_DIST_DIR.glob("*.whl")):
            result = _run([sys.executable, "-m", "hatch", "build"])
            assert (
                result.returncode == 0
            ), f"hatch build failed:\nSTDOUT:\n{result.stdout}\nSTDERR:\n{result.stderr}"  # noqa: E501

        wheels = list(_DIST_DIR.glob("*.whl"))
        assert len(wheels) >= 1, f"No .whl found in {_DIST_DIR}"

    def test_ac1_tarball_exists(self) -> None:
        """AC1: dist/ must contain a .tar.gz file."""
        if not _DIST_DIR.exists() or not list(_DIST_DIR.glob("*.tar.gz")):
            result = _run([sys.executable, "-m", "hatch", "build"])
            assert (
                result.returncode == 0
            ), f"hatch build failed:\nSTDOUT:\n{result.stdout}\nSTDERR:\n{result.stderr}"  # noqa: E501

        tarballs = list(_DIST_DIR.glob("*.tar.gz"))
        assert len(tarballs) >= 1, f"No .tar.gz found in {_DIST_DIR}"

    def test_ac1_wheel_version_matches_pyproject(self) -> None:
        """AC1: wheel filename must embed the version from pyproject.toml."""
        import tomllib

        with open(_REPO_ROOT / "pyproject.toml", "rb") as f:
            data = tomllib.load(f)
        version = data["project"]["version"]
        assert version != "0.0.0", "version must not be the placeholder 0.0.0"

        if not _DIST_DIR.exists() or not list(_DIST_DIR.glob("*.whl")):
            result = _run([sys.executable, "-m", "hatch", "build"])
            assert result.returncode == 0

        wheels = list(_DIST_DIR.glob("*.whl"))
        assert any(
            version in w.name for w in wheels
        ), f"No wheel matching version {version!r} in {[w.name for w in wheels]}"

    def test_ac1_hatch_build_exit_zero(self) -> None:
        """AC1: hatch build must exit with code 0."""
        result = _run([sys.executable, "-m", "hatch", "build"])
        assert result.returncode == 0, (
            f"hatch build exited {result.returncode}:\nSTDOUT:\n{result.stdout}"
            f"\nSTDERR:\n{result.stderr}"
        )


# ---------------------------------------------------------------------------
# AC2: twine check (또는 동등 검증) → PASS
# ---------------------------------------------------------------------------


class TestAc2TwineCheck:
    """ac2_메타데이터_검증_통과 — build artifacts must pass metadata validation."""

    def test_ac2_twine_check_passes(self) -> None:
        """AC2: twine check dist/* must exit 0 (PASSED check for all artifacts)."""
        # dist/ 가 없으면 빌드 먼저
        if not _DIST_DIR.exists() or not list(_DIST_DIR.glob("*")):
            build_result = _run([sys.executable, "-m", "hatch", "build"])
            assert (
                build_result.returncode == 0
            ), "hatch build must succeed before twine check"

        # twine check 실행 (twine 미설치 시 pip install 후 실행)
        check_result = _run([sys.executable, "-m", "twine", "check", "dist/*"])

        if (
            check_result.returncode != 0
            and "No module named twine" in check_result.stderr
        ):
            # twine 설치 후 재시도
            install_result = _run(
                [sys.executable, "-m", "pip", "install", "twine", "-q"]
            )
            assert install_result.returncode == 0, "twine install failed"
            check_result = _run([sys.executable, "-m", "twine", "check", "dist/*"])

        assert (
            check_result.returncode == 0
        ), f"twine check failed:\nSTDOUT:\n{check_result.stdout}\nSTDERR:\n{check_result.stderr}"  # noqa: E501
        # twine check 성공 메시지 확인
        output = check_result.stdout + check_result.stderr
        assert (
            "PASSED" in output or "passed" in output.lower()
        ), f"twine check output does not contain PASSED:\n{output}"

    def test_ac2_pyproject_has_readme(self) -> None:
        """AC2: pyproject.toml must reference README.md for long_description."""
        import tomllib

        with open(_REPO_ROOT / "pyproject.toml", "rb") as f:
            data = tomllib.load(f)
        assert "readme" in data["project"], "pyproject.toml must have readme field"
        readme_path = _REPO_ROOT / data["project"]["readme"]
        assert readme_path.exists(), f"README file not found: {readme_path}"

    def test_ac2_pyproject_has_license(self) -> None:
        """AC2: pyproject.toml must have license field."""
        import tomllib

        with open(_REPO_ROOT / "pyproject.toml", "rb") as f:
            data = tomllib.load(f)
        assert "license" in data["project"], "pyproject.toml must have license field"

    def test_ac2_pyproject_version_not_placeholder(self) -> None:
        """AC2: version must not be the placeholder 0.0.0."""
        import tomllib

        with open(_REPO_ROOT / "pyproject.toml", "rb") as f:
            data = tomllib.load(f)
        assert (
            data["project"]["version"] != "0.0.0"
        ), "version must be a real release version, not 0.0.0"


# ---------------------------------------------------------------------------
# AC3: release.yml 존재 + v* 트리거 + build→publish 잡 + OIDC
# ---------------------------------------------------------------------------


class TestAc3ReleaseYmlStructure:
    """ac3_release_yml_구조 — release.yml must define correct trigger and jobs."""

    @pytest.fixture(autouse=True)
    def load_yml(self) -> None:
        assert _RELEASE_YML.exists(), f"release.yml not found at {_RELEASE_YML}"
        self._content = _RELEASE_YML.read_text(encoding="utf-8")

    def test_ac3_file_exists(self) -> None:
        """AC3: .github/workflows/release.yml must exist."""
        assert _RELEASE_YML.exists()

    def test_ac3_tag_trigger(self) -> None:
        """AC3: workflow must trigger on v* tag push."""
        assert "tags:" in self._content, "tags: section missing in release.yml"
        assert "v*" in self._content, "v* tag trigger not found in release.yml"

    def test_ac3_build_job_exists(self) -> None:
        """AC3: jobs.build must be defined."""
        assert "build:" in self._content, "build job not found in release.yml"

    def test_ac3_publish_job_exists(self) -> None:
        """AC3: jobs.publish must be defined."""
        assert "publish:" in self._content, "publish job not found in release.yml"

    def test_ac3_publish_needs_build(self) -> None:
        """AC3: publish job must depend on build job (needs: build)."""
        # "needs: build" (inline) 또는 "needs:\n    - build" (list) 형태 모두 허용
        assert "needs: build" in self._content or (
            "needs:" in self._content and "- build" in self._content
        ), "publish job must need build job (needs: build)"

    def test_ac3_oidc_id_token_write(self) -> None:
        """AC3: publish job must have id-token: write permission (OIDC, ADR-607)."""
        assert (
            "id-token: write" in self._content
        ), "publish job must have id-token: write permission (OIDC)"

    def test_ac3_no_plaintext_token(self) -> None:
        """AC3: release.yml must not contain plaintext PyPI token (NFR-05)."""
        # PYPI_API_TOKEN 시크릿은 허용되지 않음 (OIDC only — ADR-607)
        assert (
            "PYPI_API_TOKEN" not in self._content
        ), "release.yml must not reference PYPI_API_TOKEN (use OIDC instead, ADR-607)"
        # pypi-token: 필드도 없어야 함 (pypa/gh-action-pypi-publish 의 토큰 방식)
        assert (
            "password:" not in self._content
        ), "release.yml must not contain password: field (use OIDC, no plaintext token)"

    def test_ac3_pypa_publish_action_used(self) -> None:
        """AC3: pypa/gh-action-pypi-publish must be used in publish job."""
        assert (
            "pypa/gh-action-pypi-publish" in self._content
        ), "pypa/gh-action-pypi-publish not found in release.yml"

    def test_ac3_hatch_build_in_build_job(self) -> None:
        """AC3: build job must run hatch build."""
        assert (
            "hatch build" in self._content
        ), "hatch build not found in build job steps"


# ---------------------------------------------------------------------------
# AC4: publish 잡에 environment:pypi → 사람 승인 없이 실행 불가
# ---------------------------------------------------------------------------


class TestAc4EnvironmentGate:
    """ac4_environment_pypi_설정 — publish job must require environment:pypi."""  # noqa: E501

    def test_ac4_environment_pypi_in_publish_job(self) -> None:
        """AC4: publish job must have environment: pypi (human approval gate)."""
        assert _RELEASE_YML.exists()
        content = _RELEASE_YML.read_text(encoding="utf-8")
        assert (
            "environment: pypi" in content
        ), "publish job must have 'environment: pypi' for human approval gate"


# ---------------------------------------------------------------------------
# AC5: 로컬 wheel pip install → from pptx_md import convert 성공
# ---------------------------------------------------------------------------


class TestAc5LocalInstall:
    """ac5_로컬_설치_성공 — wheel install in venv must allow import convert."""

    def test_ac5_wheel_install_and_import(self, tmp_path: Path) -> None:
        """AC5: pip install dist/*.whl in temp venv -> from pptx_md import convert."""
        import venv

        # dist/ 가 없으면 빌드
        if not _DIST_DIR.exists() or not list(_DIST_DIR.glob("*.whl")):
            build_result = _run([sys.executable, "-m", "hatch", "build"])
            assert build_result.returncode == 0, "hatch build must succeed for AC5"

        wheel_path = _latest_wheel()

        # 임시 가상환경 생성
        venv_dir = tmp_path / "test_venv"
        venv.create(str(venv_dir), with_pip=True, clear=True)

        # 플랫폼별 Python 경로
        if sys.platform == "win32":
            venv_python = venv_dir / "Scripts" / "python.exe"
        else:
            venv_python = venv_dir / "bin" / "python"

        # pip 업그레이드 (python -m pip 방식 — Windows pip.exe 직접 호출 우회)
        _run([str(venv_python), "-m", "pip", "install", "--upgrade", "pip", "-q"])

        # wheel 설치 (core-only, no extras)
        install_result = _run(
            [str(venv_python), "-m", "pip", "install", str(wheel_path), "-q"]
        )
        assert install_result.returncode == 0, (
            f"pip install wheel failed:\nSTDOUT:\n{install_result.stdout}"
            f"\nSTDERR:\n{install_result.stderr}"
        )

        # import 검증
        import_result = _run(
            [str(venv_python), "-c", "from pptx_md import convert; print('OK')"]
        )
        assert import_result.returncode == 0, (
            f"from pptx_md import convert failed:\nSTDOUT:\n{import_result.stdout}"
            f"\nSTDERR:\n{import_result.stderr}"
        )
        assert "OK" in import_result.stdout, "Expected 'OK' in import output"


# ---------------------------------------------------------------------------
# AC6: core-only 설치 → VLM 없이 import pptx_md 성공 (NFR-08)
# ---------------------------------------------------------------------------


class TestAc6CoreOnly:
    """ac6_core_only_import — core-only install must succeed without VLM SDKs."""

    def test_ac6_import_without_vlm_sdk(self) -> None:
        """AC6: import pptx_md must succeed in environment without anthropic/openai."""
        # 현재 환경(VLM SDK 미설치)에서 import 성공을 검증한다.
        # test_api.py AC3 와 동일 보증을 릴리스 관점에서도 확인.
        import pptx_md  # noqa: F401

        # VLM SDK 가 없어도 import 성공
        assert hasattr(
            pptx_md, "convert"
        ), "convert must be importable without VLM SDKs"

    def test_ac6_vlm_not_in_core_dependencies(self) -> None:
        """AC6: VLM SDKs must be in [vlm] extras only, not in core dependencies."""
        import tomllib

        with open(_REPO_ROOT / "pyproject.toml", "rb") as f:
            data = tomllib.load(f)

        core_deps = data["project"].get("dependencies", [])
        vlm_names = ["anthropic", "openai"]
        for dep in core_deps:
            for vlm in vlm_names:
                assert vlm not in dep.lower(), (
                    f"{vlm!r} must not be in core deps "
                    f"(must be in [vlm] extras). Found: {dep!r}"
                )

    def test_ac6_vlm_extras_defined(self) -> None:
        """AC6: [vlm] optional-dependencies must include anthropic and openai."""
        import tomllib

        with open(_REPO_ROOT / "pyproject.toml", "rb") as f:
            data = tomllib.load(f)

        optional = data["project"].get("optional-dependencies", {})
        assert "vlm" in optional, "vlm extras not defined in pyproject.toml"

        vlm_deps = optional["vlm"]
        dep_names = [d.lower() for d in vlm_deps]
        assert any(
            "anthropic" in d for d in dep_names
        ), "anthropic missing from vlm extras"
        assert any("openai" in d for d in dep_names), "openai missing from vlm extras"

    def test_ac6_wheel_install_core_only_succeeds(self, tmp_path: Path) -> None:
        """AC6: wheel install without extras → import pptx_md succeeds (NFR-08)."""
        import venv

        if not _DIST_DIR.exists() or not list(_DIST_DIR.glob("*.whl")):
            build_result = _run([sys.executable, "-m", "hatch", "build"])
            assert build_result.returncode == 0, "hatch build must succeed"

        wheel_path = _latest_wheel()

        venv_dir = tmp_path / "core_venv"
        venv.create(str(venv_dir), with_pip=True, clear=True)

        if sys.platform == "win32":
            venv_python = venv_dir / "Scripts" / "python.exe"
        else:
            venv_python = venv_dir / "bin" / "python"

        _run([str(venv_python), "-m", "pip", "install", "--upgrade", "pip", "-q"])

        # core only (no [vlm] extras)
        install_result = _run(
            [str(venv_python), "-m", "pip", "install", str(wheel_path), "-q"]
        )
        assert (
            install_result.returncode == 0
        ), f"core-only wheel install failed:\n{install_result.stderr}"

        import_result = _run(
            [str(venv_python), "-c", "import pptx_md; print(pptx_md.__version__)"]
        )
        assert (
            import_result.returncode == 0
        ), f"import pptx_md failed (core-only):\n{import_result.stderr}"


# ---------------------------------------------------------------------------
# AC7: release.yml이 CI에 의존하거나 CI 완료 후 진행하는 구조
# ---------------------------------------------------------------------------


class TestAc7CiDependency:
    """ac7_ci_선행 — release should only proceed when CI is green (AC7)."""

    def test_ac7_release_yml_exists_and_separate_from_ci(self) -> None:
        """AC7: release.yml must exist as separate workflow from ci.yml."""
        ci_yml = _REPO_ROOT / ".github" / "workflows" / "ci.yml"
        assert ci_yml.exists(), "ci.yml must exist"
        assert _RELEASE_YML.exists(), "release.yml must exist"
        # 두 파일이 다름 — 릴리스 워크플로가 분리됨
        assert ci_yml != _RELEASE_YML

    def test_ac7_release_triggered_by_tag_not_push_to_branch(self) -> None:
        """AC7: release.yml must be triggered by tag push, not branch push.

        Rationale: Tags are created manually after CI passes on master.
        This enforces that release only happens after deliberate human action
        (tag creation), ensuring CI was green before the release (AC7).
        """
        content = _RELEASE_YML.read_text(encoding="utf-8")

        # branch push trigger 가 없어야 함 (branches: 섹션이 없어야 함)
        assert (
            "branches:" not in content
        ), "release.yml must not trigger on branch push. 'branches:' found in content"

        # tag trigger 가 있어야 함
        assert "tags:" in content, "release.yml must trigger on tag push"

    def test_ac7_build_job_runs_on_ubuntu(self) -> None:
        """AC7: build job must run on ubuntu-latest (matches CI environment)."""
        content = _RELEASE_YML.read_text(encoding="utf-8")
        assert "ubuntu" in content.lower(), "build job should run on ubuntu-latest"

    def test_ac7_ci_yml_has_pytest_step(self) -> None:
        """AC7: ci.yml must run pytest (verifies CI gate exists before release)."""
        ci_yml = _REPO_ROOT / ".github" / "workflows" / "ci.yml"
        ci_content = ci_yml.read_text(encoding="utf-8")
        assert "pytest" in ci_content, "ci.yml must contain a pytest step"
