import subprocess, os, sys, platform

def venv_exists():
    return os.path.exists("venv")

def create_venv():
    print("🔧 Creating virtual environment...")
    subprocess.check_call([sys.executable, "-m", "venv", "venv"])

def install_requirements():
    print("📦 Installing dependencies...")
    pip_exec = os.path.join("venv", "Scripts" if platform.system() == "Windows" else "bin", "pip")
    subprocess.check_call([pip_exec, "install", "-r", "requirements.txt"])

def run_script():
    print("🚀 Running JSON Generator in virtual environment...")
    python_exec = os.path.join("venv", "Scripts" if platform.system() == "Windows" else "bin", "python")
    subprocess.check_call([python_exec, "JsonGenerator.py"])

if __name__ == "__main__":
    if not venv_exists():
        create_venv()
        install_requirements()
    run_script()
