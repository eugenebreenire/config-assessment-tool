#!/usr/bin/env python
"""
One-click launcher for the Config Assessment Tool.

- Creates a .venv if it doesn't exist
- Installs requirements
- Starts the Flask UI (webapp.app)
"""

import os
import subprocess
from pathlib import Path


def main():
    # Define the virtual environment path
    venv_path = Path(".venv")
    venv_python = venv_path / "bin" / "python"
    venv_pip = venv_path / "bin" / "pip"

    # Step 1: Create the virtual environment if it doesn't exist
    if not venv_path.exists():
        print("Creating virtual environment...")
        subprocess.check_call(["python3", "-m", "venv", str(venv_path)])

    # Step 2: Install dependencies
    print("Installing dependencies...")
    subprocess.check_call([str(venv_pip), "install", "-r", "requirements.txt"])

    # Step 3: Run the Flask application
    print("Starting the application...")
    subprocess.check_call([str(venv_python), "-m", "webapp.app"])

if __name__ == "__main__":
    main()
