#!/usr/bin/env python3
"""Clear Python cache files to ensure clean code reload."""

import os
import shutil
import glob

def clear_python_cache():
    """Clear all Python cache files and directories."""
    print("🧹 Clearing Python cache files...")
    
    # Clear __pycache__ directories
    pycache_dirs = glob.glob("**/__pycache__", recursive=True)
    for pycache_dir in pycache_dirs:
        try:
            shutil.rmtree(pycache_dir)
            print(f"  ✅ Removed {pycache_dir}")
        except Exception as e:
            print(f"  ⚠️  Failed to remove {pycache_dir}: {e}")
    
    # Clear .pyc files
    pyc_files = glob.glob("**/*.pyc", recursive=True)
    for pyc_file in pyc_files:
        try:
            os.remove(pyc_file)
            print(f"  ✅ Removed {pyc_file}")
        except Exception as e:
            print(f"  ⚠️  Failed to remove {pyc_file}: {e}")
    
    # Clear .pyo files
    pyo_files = glob.glob("**/*.pyo", recursive=True)
    for pyo_file in pyo_files:
        try:
            os.remove(pyo_file)
            print(f"  ✅ Removed {pyo_file}")
        except Exception as e:
            print(f"  ⚠️  Failed to remove {pyo_file}: {e}")
    
    print("🎉 Cache clearing complete!")
    print("Now restart the server to ensure fresh code loading.")

if __name__ == "__main__":
    clear_python_cache()