from pathlib import Path
import shutil
import sys

ROOT = Path(__file__).resolve().parent
PAYLOAD = ROOT / "payload"
VERSION = "0.7.1"

SOURCE = PAYLOAD / "word_page.py"
TARGET = ROOT / "word_page.py"
VERSION_FILE = ROOT / "version.txt"

def fail(message):
    print(f"\nERROR: {message}\n")
    sys.exit(1)

if not SOURCE.exists():
    fail("Missing payload/word_page.py")

if not TARGET.exists():
    fail(
        "word_page.py was not found. "
        "Extract this installer inside your ToolPy folder."
    )

shutil.copy2(
    TARGET,
    ROOT / "word_page.py.before_version_display.bak",
)

if VERSION_FILE.exists():
    shutil.copy2(
        VERSION_FILE,
        ROOT / "version.txt.before_version_display.bak",
    )

shutil.copy2(SOURCE, TARGET)
VERSION_FILE.write_text(
    VERSION + "\n",
    encoding="utf-8",
)

# Remove the payload after a successful install.
shutil.rmtree(PAYLOAD, ignore_errors=True)

readme = ROOT / "README.txt"
if readme.exists():
    readme.unlink()

print("\nVersion display installed successfully.")
print(f"version.txt updated automatically to {VERSION}")
print("Updated: word_page.py")
print("\nThe GUI will show:")
print(f"  ToolPy v{VERSION}")
print("\nRebuild with:")
print("  py -m PyInstaller --clean ToolPy.spec")
