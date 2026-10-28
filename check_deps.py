"""Install missing packages on every launch — no one-time marker, so a
package added in a later version is never silently skipped."""
import importlib.util
import re
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
REQUIREMENTS = ROOT / "requirements.txt"

IMPORT_NAME = {
    "pillow": "PIL",
    "opencv-python-headless": "cv2",
}


def import_name(pip_name: str) -> str:
    return IMPORT_NAME.get(pip_name.lower(), pip_name.replace("-", "_"))


def required_packages():
    names = []
    for line in REQUIREMENTS.read_text().splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        match = re.match(r"^([A-Za-z0-9_.\-]+)", line)
        if match:
            names.append(match.group(1))
    return names


def missing_packages():
    missing = []
    for pip_name in required_packages():
        if importlib.util.find_spec(import_name(pip_name)) is None:
            missing.append(pip_name)
    return missing


def requirements_fingerprint() -> str:
    return str(hash(REQUIREMENTS.read_text()))


def main() -> int:
    verify_only = "--verify" in sys.argv
    missing = missing_packages()
    if not missing:
        print("All dependencies present.")
        return 0
    if verify_only:
        print(f"Still missing: {missing}")
        return 1
    print(f"Installing: {missing}")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", str(REQUIREMENTS)])
    return 0


if __name__ == "__main__":
    sys.exit(main())
