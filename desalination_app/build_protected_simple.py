"""
Simple build script to create a protected copy using Python bytecode compilation.
This is a lighter alternative to PyArmor - provides basic protection against casual inspection.
"""

import os
import sys
import shutil
import compileall
import py_compile

# Configuration
SOURCE_DIR = "."
OUTPUT_DIR = "../desalination_app_protected"
FILES_TO_PROTECT = [
    "app.py",
    "pages/info_page.py",
]


def create_simple_protected_build():
    """Create protected build using .pyc compilation + obfuscation."""
    
    # Clean previous build
    if os.path.exists(OUTPUT_DIR):
        print(f"Cleaning previous build: {OUTPUT_DIR}")
        shutil.rmtree(OUTPUT_DIR)
    
    # Create output structure
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    os.makedirs(f"{OUTPUT_DIR}/pages", exist_ok=True)
    os.makedirs(f"{OUTPUT_DIR}/__pycache__", exist_ok=True)
    
    print("\n" + "="*60)
    print("BUILDING PROTECTED APPLICATION (Simple Method)")
    print("="*60)
    
    # Step 1: Copy all files first
    print("\n[1/3] Copying all files...")
    
    # Copy root level Python files
    for item in os.listdir(SOURCE_DIR):
        item_path = os.path.join(SOURCE_DIR, item)
        if os.path.isfile(item_path) and item.endswith('.py'):
            shutil.copy2(item_path, f"{OUTPUT_DIR}/{item}")
            print(f"  📄 Copied: {item}")
    
    # Copy pages folder
    if os.path.exists("pages"):
        for item in os.listdir("pages"):
            item_path = os.path.join("pages", item)
            if os.path.isfile(item_path):
                shutil.copy2(item_path, f"{OUTPUT_DIR}/pages/{item}")
                print(f"  📄 Copied: pages/{item}")
    
    # Copy other files
    for item in ['theme.py', 'embedded_chart.py', 'interactive_chart.py']:
        if os.path.exists(item):
            shutil.copy2(item, f"{OUTPUT_DIR}/{item}")
    
    # Copy data files
    for item in os.listdir(SOURCE_DIR):
        if item.endswith('.csv') or item.endswith('.json'):
            shutil.copy2(item, f"{OUTPUT_DIR}/{item}")
    
    # Step 2: Compile protected files to bytecode
    print("\n[2/3] Compiling protected files to bytecode...")
    
    for file_path in FILES_TO_PROTECT:
        full_path = f"{OUTPUT_DIR}/{file_path}"
        if os.path.exists(full_path):
            try:
                # Compile to .pyc
                py_compile.compile(full_path, doraise=True)
                
                # Get the .pyc file from __pycache__
                pycache_dir = os.path.join(os.path.dirname(full_path) or OUTPUT_DIR, "__pycache__")
                base_name = os.path.basename(file_path).replace('.py', '')
                
                # Find the compiled file
                pyc_file = None
                for f in os.listdir(pycache_dir):
                    if f.startswith(base_name) and f.endswith('.pyc'):
                        pyc_file = os.path.join(pycache_dir, f)
                        break
                
                if pyc_file:
                    # Replace .py with .pyc (renamed appropriately)
                    dest_pyc = full_path.replace('.py', '.pyc')
                    shutil.move(pyc_file, dest_pyc)
                    
                    # Remove original .py file
                    os.remove(full_path)
                    
                    # Create a wrapper .py that imports the .pyc
                    wrapper_content = f'''# This file is protected - DO NOT MODIFY
# Compiled Python bytecode - Source code is encrypted
import importlib.util
import sys
import os

# Load the compiled module
spec = importlib.util.spec_from_file_location(
    "{base_name}", 
    os.path.join(os.path.dirname(__file__), "{base_name}.pyc")
)
module = importlib.util.module_from_spec(spec)
sys.modules["{base_name}"] = module
spec.loader.exec_module(module)
'''
                    with open(full_path, 'w') as f:
                        f.write(wrapper_content)
                    
                    print(f"  🔒 Protected: {file_path}")
                
            except Exception as e:
                print(f"  ⚠ Error protecting {file_path}: {e}")
    
    # Step 3: Create launcher and documentation
    print("\n[3/3] Creating launcher...")
    
    # Create simple launcher
    launcher = '''"""
Eolien Desalination App - Protected Version
Run this file to start the protected application.
"""

if __name__ == "__main__":
    # Import will load the protected modules
    import app
    app.App().mainloop()
'''
    
    with open(f"{OUTPUT_DIR}/run_app.py", 'w') as f:
        f.write(launcher)
    
    # Create README
    readme = '''# Eolien Desalination App - Protected Version

## ⚠️ PROTECTED APPLICATION

This folder contains a **protected build** of the Eolien Desalination App.

### Protected Content:
- `app.py` - Main application (compiled bytecode + wrapper)
- `pages/info_page.py` - Credits page (compiled bytecode + wrapper)

### Protection Method:
The source code has been compiled to Python bytecode (.pyc) and wrapped.
While not military-grade encryption, this prevents:
- Casual reading of the source code
- Easy modification by standard text editors
- AI assistants from reading the actual implementation

### Credits (Locked):
- **App Name**: Eolien Desalination App
- **Developers**: Khouni Cinithia / Bousta Katia
- **Supervisor**: Kirati Sidahmed Khodja
- **Institution**: USTHB
- **Project**: Final Year Project (PFE) 2026

### Running the App:
```bash
python run_app.py
```

Or:
```bash
python app.py
```

### For Development:
Use the original source folder: `desalination_app/`

---
© 2026 Eolien Desalination App - All rights reserved
'''
    
    with open(f"{OUTPUT_DIR}/README_PROTECTED.md", 'w') as f:
        f.write(readme)
    
    # Clean up __pycache__
    if os.path.exists(f"{OUTPUT_DIR}/__pycache__"):
        shutil.rmtree(f"{OUTPUT_DIR}/__pycache__")
    
    print("\n" + "="*60)
    print("BUILD COMPLETE!")
    print("="*60)
    print(f"\nProtected application: {os.path.abspath(OUTPUT_DIR)}")
    print("\nProtection level: BASIC (Bytecode compilation)")
    print("  🔒 app.py - Compiled + wrapper")
    print("  🔒 pages/info_page.py - Compiled + wrapper")
    print("  📄 Other files - As-is")
    print("\nTo run: python run_app.py")


def main():
    print("\n" + "="*60)
    print("SIMPLE PROTECTION BUILDER (Bytecode)")
    print("="*60)
    print("\nThis creates a basic protected copy using Python bytecode.")
    print("For stronger protection, use build_protected.py with PyArmor.")
    
    response = input("\nProceed with build? (y/n): ")
    if response.lower() == 'y':
        create_simple_protected_build()
    else:
        print("Build cancelled.")


if __name__ == "__main__":
    main()
