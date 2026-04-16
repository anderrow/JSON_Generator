import subprocess
import sys

from run_env import (
    PROJECT_DIR,
    create_venv,
    get_venv_python,
    install_requirements,
    venv_exists,
    warn_about_in_repo_venv,
)


def run_sync():
    subprocess.check_call(
        [str(get_venv_python()), str(PROJECT_DIR / "sync_processview.py"), *sys.argv[1:]],
        cwd=PROJECT_DIR,
    )


if __name__ == "__main__":
    warn_about_in_repo_venv()

    if not venv_exists():
        create_venv()
        install_requirements()

    run_sync()
