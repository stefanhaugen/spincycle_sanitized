"""
Bootstrap pip into the embedded Python installation.
This script is called by SETUP_once.bat — do not run manually.
It adds the bundled pip wheel to sys.path and installs pip + setuptools
into the embedded Python's Lib/site-packages.
"""
import sys
import os

app_dir = os.path.dirname(os.path.abspath(__file__))
packages_dir = os.path.join(app_dir, "packages")

# Find the pip wheel in the packages folder
pip_wheel = None
for f in os.listdir(packages_dir):
    if f.startswith("pip-") and f.endswith(".whl"):
        pip_wheel = os.path.join(packages_dir, f)
        break

if pip_wheel is None:
    print("[ERROR] Could not find pip wheel in packages folder.")
    sys.exit(1)

# Add the wheel to sys.path so we can import pip
sys.path.insert(0, pip_wheel)

# Run pip install for pip + setuptools from the local wheel cache
from pip._internal.cli.main import main as pip_main
sys.exit(pip_main([
    "install",
    "--no-index",
    "--find-links", packages_dir,
    "pip", "setuptools",
]))
