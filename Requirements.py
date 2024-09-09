import subprocess
import sys

# List of required packages
required_packages = [
    'tkinter',  # Built-in, no need to install via pip
    'Pillow',  # For handling images
    'pandas',  # For data manipulation
    'tkcalendar',  # For DateEntry widget
]


def install(package):
    """Install the package using pip."""
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])


for package in required_packages:
    try:
        install(package)
        print(f"'{package}' installed successfully.")
    except subprocess.CalledProcessError as e:
        print(f"Failed to install '{package}'. Error: {e}")
