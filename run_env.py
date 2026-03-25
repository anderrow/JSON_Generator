import hashlib
import os
import platform
import subprocess
import sys
from pathlib import Path


PROJECT_DIR = Path(__file__).resolve().parent
IN_REPO_VENV_DIR = PROJECT_DIR / "venv"


def get_venv_dir():
    project_hash = hashlib.sha1(str(PROJECT_DIR).encode("utf-8")).hexdigest()[:12]

    if platform.system() == "Windows":
        base_dir = Path(os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local"))
    else:
        base_dir = Path(os.environ.get("XDG_DATA_HOME", Path.home() / ".local" / "share"))

    return base_dir / "PythonProjectVenvs" / f"{PROJECT_DIR.name}-{project_hash}"


VENV_DIR = get_venv_dir()


def get_venv_python():
    if platform.system() == "Windows":
        return VENV_DIR / "Scripts" / "python.exe"
    return VENV_DIR / "bin" / "python"


def venv_exists():
    return get_venv_python().exists()


def create_venv():
    print(f"Creating virtual environment in: {VENV_DIR}")
    VENV_DIR.parent.mkdir(parents=True, exist_ok=True)
    subprocess.check_call([sys.executable, "-m", "venv", str(VENV_DIR)], cwd=PROJECT_DIR)


def install_requirements():
    print("Installing dependencies...")
    subprocess.check_call(
        [str(get_venv_python()), "-m", "pip", "install", "-r", str(PROJECT_DIR / "requirements.txt")],
        cwd=PROJECT_DIR,
    )


def run_script():
    print("Running Excel Generator in virtual environment...")
    subprocess.check_call([str(get_venv_python()), str(PROJECT_DIR / "ExcelGenerator.py")], cwd=PROJECT_DIR)


def warn_about_in_repo_venv():
    if IN_REPO_VENV_DIR.exists():
        print(
            "Warning: found an old 'venv' folder inside the project. "
            "run_env.py no longer uses it, so you can delete it to stop Nextcloud from syncing it."
        )


if __name__ == "__main__":
    warn_about_in_repo_venv()

    if not venv_exists():
        create_venv()
        install_requirements()

    run_script()
