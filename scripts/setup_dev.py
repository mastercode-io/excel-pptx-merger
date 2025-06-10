#!/usr/bin/env python3
"""Development environment setup script."""

import os
import sys
import subprocess
import shutil
from pathlib import Path

def run_command(command, description=""):
    """Run a shell command and handle errors."""
    print(f"Running: {command}")
    if description:
        print(f"  {description}")
    
    try:
        result = subprocess.run(command, shell=True, check=True, capture_output=True, text=True)
        if result.stdout:
            print(f"  Output: {result.stdout.strip()}")
        return True
    except subprocess.CalledProcessError as e:
        print(f"  Error: {e}")
        if e.stderr:
            print(f"  Stderr: {e.stderr.strip()}")
        return False

def check_python_version():
    """Check if Python version is compatible."""
    if sys.version_info < (3, 9):
        print("Error: Python 3.9 or higher is required")
        sys.exit(1)
    print(f"✓ Python version: {sys.version}")

def check_uv_installed():
    """Check if uv package manager is installed."""
    if shutil.which('uv') is None:
        print("Installing uv package manager...")
        if not run_command("pip install uv", "Installing uv package manager"):
            print("Error: Failed to install uv")
            sys.exit(1)
    else:
        print("✓ uv package manager is installed")

def setup_virtual_environment():
    """Set up virtual environment using uv."""
    print("\nSetting up virtual environment...")
    
    if not run_command("uv venv", "Creating virtual environment"):
        print("Error: Failed to create virtual environment")
        sys.exit(1)
    
    print("✓ Virtual environment created")

def install_dependencies():
    """Install project dependencies."""
    print("\nInstalling dependencies...")
    
    # Install main dependencies
    if not run_command("uv pip install -e .", "Installing main dependencies"):
        print("Error: Failed to install main dependencies")
        sys.exit(1)
    
    # Install development dependencies
    if not run_command("uv pip install -e .[dev]", "Installing development dependencies"):
        print("Error: Failed to install development dependencies")
        sys.exit(1)
    
    print("✓ Dependencies installed")

def setup_environment_files():
    """Set up environment configuration files."""
    print("\nSetting up environment files...")
    
    env_file = Path(".env")
    env_example = Path(".env.example")
    
    if not env_file.exists() and env_example.exists():
        shutil.copy(env_example, env_file)
        print(f"✓ Created {env_file} from {env_example}")
        print(f"  Please edit {env_file} with your configuration")
    elif env_file.exists():
        print(f"✓ Environment file {env_file} already exists")
    else:
        print(f"⚠ Warning: No environment template found")

def create_temp_directories():
    """Create necessary temporary directories."""
    print("\nCreating temporary directories...")
    
    temp_dirs = [
        "temp",
        "logs",
        "/tmp/excel_pptx_merger_dev"
    ]
    
    for temp_dir in temp_dirs:
        try:
            Path(temp_dir).mkdir(parents=True, exist_ok=True)
            print(f"✓ Created directory: {temp_dir}")
        except PermissionError:
            print(f"⚠ Warning: Could not create {temp_dir} (permission denied)")

def run_tests():
    """Run test suite to verify setup."""
    print("\nRunning test suite...")
    
    if not run_command("uv run pytest tests/ -v --tb=short", "Running tests"):
        print("⚠ Warning: Some tests failed - check your setup")
        return False
    
    print("✓ All tests passed")
    return True

def setup_pre_commit_hooks():
    """Set up pre-commit hooks for code quality."""
    print("\nSetting up pre-commit hooks...")
    
    # Check if pre-commit is available
    if shutil.which('pre-commit') is None:
        if not run_command("uv pip install pre-commit", "Installing pre-commit"):
            print("⚠ Warning: Could not install pre-commit")
            return False
    
    # Create pre-commit configuration
    pre_commit_config = """
repos:
  - repo: https://github.com/psf/black
    rev: 23.7.0
    hooks:
      - id: black
        language_version: python3

  - repo: https://github.com/pycqa/flake8
    rev: 6.0.0
    hooks:
      - id: flake8

  - repo: https://github.com/pre-commit/mirrors-mypy
    rev: v1.5.0
    hooks:
      - id: mypy
        additional_dependencies: [types-requests, types-Pillow]

  - repo: https://github.com/pre-commit/pre-commit-hooks
    rev: v4.4.0
    hooks:
      - id: trailing-whitespace
      - id: end-of-file-fixer
      - id: check-yaml
      - id: check-json
"""
    
    pre_commit_file = Path(".pre-commit-config.yaml")
    if not pre_commit_file.exists():
        with open(pre_commit_file, 'w') as f:
            f.write(pre_commit_config.strip())
        print("✓ Created .pre-commit-config.yaml")
    
    if not run_command("pre-commit install", "Installing pre-commit hooks"):
        print("⚠ Warning: Could not install pre-commit hooks")
        return False
    
    print("✓ Pre-commit hooks installed")
    return True

def print_next_steps():
    """Print next steps for the developer."""
    print("\n" + "="*60)
    print("🎉 Development environment setup complete!")
    print("="*60)
    print("\nNext steps:")
    print("1. Activate your virtual environment:")
    print("   source .venv/bin/activate  # On Unix/macOS")
    print("   .venv\\Scripts\\activate     # On Windows")
    print("\n2. Edit .env file with your configuration")
    print("\n3. Start the development server:")
    print("   uv run python -m src.main serve --debug")
    print("\n4. Or run with Docker:")
    print("   docker-compose -f docker/docker-compose.yml up")
    print("\n5. Access the API at: http://localhost:5000")
    print("\n6. View API health: http://localhost:5000/api/v1/health")
    print("\nUseful commands:")
    print("  uv run pytest                 # Run tests")
    print("  uv run black src/ tests/      # Format code")
    print("  uv run flake8 src/ tests/     # Check style")
    print("  uv run mypy src/              # Type checking")

def main():
    """Main setup function."""
    print("Excel to PowerPoint Merger - Development Setup")
    print("=" * 50)
    
    # Change to project root directory
    script_dir = Path(__file__).parent
    project_root = script_dir.parent
    os.chdir(project_root)
    print(f"Working directory: {project_root}")
    
    # Run setup steps
    check_python_version()
    check_uv_installed()
    setup_virtual_environment()
    install_dependencies()
    setup_environment_files()
    create_temp_directories()
    
    # Optional steps
    tests_passed = run_tests()
    pre_commit_setup = setup_pre_commit_hooks()
    
    # Summary
    print_next_steps()
    
    if not tests_passed:
        print("\n⚠ Note: Some tests failed. Check your setup before proceeding.")
    
    if not pre_commit_setup:
        print("\n⚠ Note: Pre-commit hooks setup failed. Code quality checks may not work.")

if __name__ == "__main__":
    main()