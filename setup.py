#!/usr/bin/env python3
import subprocess
import sys
import os

def install_system_deps():
    """Install system dependencies"""
    try:
        # Update package list
        subprocess.run(['apt-get', 'update'], check=True)
        
        # Install tesseract and poppler
        subprocess.run([
            'apt-get', 'install', '-y', 
            'tesseract-ocr', 
            'poppler-utils'
        ], check=True)
        
        print("✅ System dependencies installed successfully")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ Failed to install system dependencies: {e}")
        return False
    except FileNotFoundError:
        print("⚠️ apt-get not found, assuming dependencies are already installed")
        return True

def install_python_deps():
    """Install Python dependencies"""
    try:
        subprocess.run([
            sys.executable, '-m', 'pip', 'install', 
            '-r', 'requirements.txt'
        ], check=True)
        
        print("✅ Python dependencies installed successfully")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ Failed to install Python dependencies: {e}")
        return False

if __name__ == "__main__":
    print("🚀 Starting setup...")
    
    success = True
    success &= install_system_deps()
    success &= install_python_deps()
    
    if success:
        print("✅ Setup completed successfully!")
        sys.exit(0)
    else:
        print("❌ Setup failed!")
        sys.exit(1)