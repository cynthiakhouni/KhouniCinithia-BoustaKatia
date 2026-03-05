"""
Build script to create a PyArmor-protected copy of the Eolien Desalination App.
This encrypts the code so it cannot be easily modified or inspected.
"""

import os
import sys
import shutil
import subprocess

# Configuration
SOURCE_DIR = "."
OUTPUT_DIR = "../desalination_app_protected"
FILES_TO_PROTECT = [
    "app.py",
    "pages/info_page.py",
]
FILES_TO_COPY_AS_IS = [
    "theme.py",
    "requirements.txt",
    "pages/__init__.py",
    "pages/source_page.py",
    "pages/desal_page.py",
    "pages/econ_page.py",
    "embedded_chart.py",
    "interactive_chart.py",
]


def check_pyarmor():
    """Check if PyArmor is installed."""
    try:
        import pyarmor
        return True
    except ImportError:
        return False


def install_pyarmor():
    """Install PyArmor if not present."""
    print("PyArmor not found. Installing...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pyarmor"])
    print("PyArmor installed successfully!")


def create_protected_build():
    """Create the protected build using PyArmor."""
    
    # Clean previous build
    if os.path.exists(OUTPUT_DIR):
        print(f"Cleaning previous build: {OUTPUT_DIR}")
        shutil.rmtree(OUTPUT_DIR)
    
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    os.makedirs(f"{OUTPUT_DIR}/pages", exist_ok=True)
    
    print("\n" + "="*60)
    print("BUILDING PROTECTED APPLICATION")
    print("="*60)
    
    # Step 1: Obfuscate protected files
    print("\n[1/4] Encrypting protected files...")
    for file_path in FILES_TO_PROTECT:
        if not os.path.exists(file_path):
            print(f"  ⚠ Warning: {file_path} not found, skipping...")
            continue
        
        print(f"  🔒 Protecting: {file_path}")
        
        # Create obfuscated version
        output_file = f"{OUTPUT_DIR}/{file_path}"
        
        # Use pyarmor to obfuscate
        cmd = [
            sys.executable, "-m", "pyarmor", "gen",
            "--output", os.path.dirname(output_file) or OUTPUT_DIR,
            "--restrict",  # Restrict mode - cannot be imported by plain scripts
            file_path
        ]
        
        try:
            subprocess.check_call(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            print(f"     ✓ Encrypted: {file_path}")
        except subprocess.CalledProcessError as e:
            print(f"     ✗ Failed to encrypt {file_path}: {e}")
            # Fallback: copy as-is with warning
            shutil.copy2(file_path, output_file)
            print(f"     ⚠ Copied as-is (not encrypted)")
    
    # Step 2: Copy non-protected files
    print("\n[2/4] Copying standard files...")
    for file_path in FILES_TO_COPY_AS_IS:
        if os.path.exists(file_path):
            dest = f"{OUTPUT_DIR}/{file_path}"
            os.makedirs(os.path.dirname(dest), exist_ok=True)
            shutil.copy2(file_path, dest)
            print(f"  📄 Copied: {file_path}")
        else:
            print(f"  ⚠ Not found: {file_path}")
    
    # Step 3: Copy CSV/data files if they exist
    print("\n[3/4] Copying data files...")
    for item in os.listdir(SOURCE_DIR):
        if item.endswith('.csv') or item.endswith('.json') or item.endswith('.txt'):
            if os.path.isfile(item):
                shutil.copy2(item, f"{OUTPUT_DIR}/{item}")
                print(f"  📊 Copied: {item}")
    
    # Step 4: Create a launcher script
    print("\n[4/4] Creating launcher...")
    launcher_content = '''"""
Eolien Desalination App - Protected Version
This is the encrypted/obfuscated version of the application.
The info_page.py and app.py files are protected and cannot be modified.

IMPORTANT: This folder contains encrypted code. Do not modify files here.
For development, use the original source folder.
"""

import sys
import os

# Ensure pyarmor runtime is available
try:
    from pyarmor_runtime import __pyarmor__
except ImportError:
    pass  # Will be handled by obfuscated code

# Run the application
if __name__ == "__main__":
    import app
    app.App().mainloop()
'''
    
    with open(f"{OUTPUT_DIR}/run_app.py", 'w') as f:
        f.write(launcher_content)
    
    # Create a README for the protected version
    readme_content = '''# Eolien Desalination App - Protected Version

## ⚠️ IMPORTANT NOTICE

This folder contains **ENCRYPTED and OBFUSCATED** code protected by PyArmor.

### Protected Files (Cannot be modified):
- `app.py` - Main application (encrypted)
- `pages/info_page.py` - Credits page (encrypted)

### What This Means:
- ✅ The application will run normally
- ✅ Users cannot read or modify the source code
- ✅ The credits (names, institution, year) are locked
- ❌ Do NOT attempt to edit files in this folder
- ❌ AI assistants cannot read or modify the encrypted code

### For Development:
Use the original source folder: `desalination_app/`

### How to Run:
```bash
python run_app.py
```

Or directly:
```bash
python app.py
```

### Credits (Locked):
- **App Name**: Eolien Desalination App
- **Developers**: Khouni Cinithia / Bousta Katia
- **Supervisor**: Kirati Sidahmed Khodja
- **Institution**: USTHB
- **Project**: Final Year Project (PFE) 2026

---
Protected with PyArmor - All rights reserved © 2026
'''
    
    with open(f"{OUTPUT_DIR}/README_PROTECTED.md", 'w') as f:
        f.write(readme_content)
    
    print("\n" + "="*60)
    print("BUILD COMPLETE!")
    print("="*60)
    print(f"\nProtected application created in: {os.path.abspath(OUTPUT_DIR)}")
    print("\nKey features:")
    print("  🔒 app.py - ENCRYPTED")
    print("  🔒 pages/info_page.py - ENCRYPTED")
    print("  📄 Other files - Copied as-is")
    print("  🚀 Run with: python run_app.py")
    print("\n⚠️  Do NOT edit files in the protected folder!")
    print("    Use the original folder for development.")


def main():
    """Main entry point."""
    print("\n" + "="*60)
    print("PYARMOR PROTECTION BUILDER")
    print("="*60)
    
    # Check/install PyArmor
    if not check_pyarmor():
        response = input("PyArmor is not installed. Install now? (y/n): ")
        if response.lower() == 'y':
            install_pyarmor()
        else:
            print("PyArmor is required. Exiting.")
            sys.exit(1)
    
    # Confirm build
    print(f"\nThis will create a protected copy at:")
    print(f"  {os.path.abspath(OUTPUT_DIR)}")
    print(f"\nFiles to encrypt: {', '.join(FILES_TO_PROTECT)}")
    response = input("\nProceed? (y/n): ")
    
    if response.lower() == 'y':
        create_protected_build()
    else:
        print("Build cancelled.")


if __name__ == "__main__":
    main()
