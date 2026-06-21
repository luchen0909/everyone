from __future__ import annotations

import py_compile
import shutil
import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
APP_DIR = ROOT / "项目1"
BALANCE_SRC = ROOT / "静态平衡表" / "src"
APP_NAME = "SupplyCoordinationToolV3"
SPEC_NAME = "供需协同工具V3.0_macOS.spec"


def run(cmd: list[str], cwd: Path | None = None) -> None:
    print("+", " ".join(cmd), flush=True)
    subprocess.run(cmd, cwd=str(cwd) if cwd else None, check=True)


def compile_sources() -> None:
    for path in [
        APP_DIR / "BOM查询工具.py",
        APP_DIR / "bom_legacy_tk.py",
        BALANCE_SRC / "mrp_balance_tool" / "pipeline.py",
        BALANCE_SRC / "mrp_balance_tool" / "gui.py",
    ]:
        if not path.exists():
            raise FileNotFoundError(path)
        py_compile.compile(str(path), doraise=True)


def main() -> None:
    requirements = APP_DIR / "requirements-macos.txt"
    template = APP_DIR / "assets" / "静态平衡表模板.xlsx"
    spec = APP_DIR / SPEC_NAME
    app_path = APP_DIR / "dist" / f"{APP_NAME}.app"
    exe_path = app_path / "Contents" / "MacOS" / APP_NAME
    artifact_dir = ROOT / "artifact"
    artifact_zip = artifact_dir / "supply-coordination-macos.zip"

    for path in [requirements, template, spec]:
        if not path.exists():
            raise FileNotFoundError(path)
    if template.stat().st_size <= 0:
        raise RuntimeError(f"Template is empty: {template}")

    run([sys.executable, "-m", "pip", "install", "--upgrade", "pip", "setuptools", "wheel"])
    run([sys.executable, "-m", "pip", "install", "-r", str(requirements)])
    compile_sources()
    run([sys.executable, "-m", "PyInstaller", "--clean", "--noconfirm", str(spec)], cwd=APP_DIR)

    if not app_path.exists():
        raise FileNotFoundError(app_path)
    if not exe_path.exists():
        raise FileNotFoundError(exe_path)

    run(["codesign", "--force", "--deep", "--sign", "-", str(app_path)])
    run(["codesign", "--verify", "--deep", "--strict", "--verbose=2", str(app_path)])
    run(["plutil", "-lint", str(app_path / "Contents" / "Info.plist")])

    if artifact_dir.exists():
        shutil.rmtree(artifact_dir)
    artifact_dir.mkdir(parents=True)
    run(["ditto", "-c", "-k", "--sequesterRsrc", "--keepParent", str(app_path), str(artifact_zip)])

    if not artifact_zip.exists() or artifact_zip.stat().st_size <= 0:
        raise RuntimeError(f"Artifact was not created: {artifact_zip}")
    print(f"Created artifact: {artifact_zip}")


if __name__ == "__main__":
    main()
