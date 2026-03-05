# Application Protection Guide

This guide explains how to create a protected copy of your Eolien Desalination App that prevents modification of the credits/info page.

## Overview

You have **two options** for protection:

### Option 1: Strong Protection (PyArmor) - RECOMMENDED
**Best for:** Maximum security, production distribution

PyArmor encrypts your Python code at the bytecode level with runtime encryption.

#### How to use:
```bash
# Install PyArmor (one time)
pip install pyarmor

# Run the build script
python build_protected.py
```

#### What it does:
- Encrypts `app.py` and `pages/info_page.py` with AES encryption
- Runtime protection prevents debugging/inspection
- Code cannot be read by AI or humans
- Files will be created in: `../desalination_app_protected/`

#### Protection Level: 🔒🔒🔒 HIGH
- Source code is completely unreadable
- Cannot be modified without the original source
- Runtime decryption in memory only

---

### Option 2: Basic Protection (Bytecode)
**Best for:** Quick protection, internal use

Compiles Python to bytecode (.pyc) with wrapper scripts.

#### How to use:
```bash
# No installation needed - uses Python built-in tools
python build_protected_simple.py
```

#### What it does:
- Compiles protected files to bytecode
- Removes original source code
- Creates wrapper that loads bytecode
- Files will be created in: `../desalination_app_protected/`

#### Protection Level: 🔒 MEDIUM
- Source code is not human-readable
- Can be reverse-engineered by determined attackers
- Sufficient to prevent casual modification

---

## Folder Structure After Protection

### Original (Development):
```
desalination_app/
├── app.py                    ← Editable
├── pages/
│   ├── info_page.py          ← Editable
│   └── ...
├── build_protected.py
└── ...
```

### Protected (Distribution):
```
desalination_app_protected/
├── app.py                    ← Encrypted/Protected
├── app.pyc                   ← Compiled bytecode (if using simple method)
├── pages/
│   ├── info_page.py          ← Encrypted/Protected
│   └── ...
├── run_app.py                ← Launcher
└── README_PROTECTED.md
```

---

## Protected Information

The following information will be locked and cannot be changed in the protected version:

| Field | Value |
|-------|-------|
| **App Name** | Eolien Desalination App |
| **Developers** | Khouni Cinithia / Bousta Katia |
| **Supervisor** | Kirati Sidahmed Khodja |
| **Institution** | USTHB |
| **Project** | Final Year Project (PFE) 2026 |

---

## Running the Protected App

After building, navigate to the protected folder:

```bash
cd ../desalination_app_protected
python run_app.py
```

Or:
```bash
cd ../desalination_app_protected
python app.py
```

The application will work exactly the same - users just cannot see or modify the protected code.

---

## Which Option Should You Choose?

| Scenario | Recommendation |
|----------|---------------|
| Final project submission | **Option 1 (PyArmor)** |
| Sharing with reviewers | **Option 1 (PyArmor)** |
| Quick internal test | **Option 2 (Bytecode)** |
| Maximum security | **Option 1 (PyArmor)** |

---

## Important Notes

1. **Original folder remains editable** - Keep developing in `desalination_app/`
2. **Protected folder is read-only** - Never edit files in `desalination_app_protected/`
3. **Rebuild after changes** - If you update the original, run the build script again
4. **PyArmor requires installation** - But provides much stronger protection

---

## Troubleshooting

### "PyArmor not found"
```bash
pip install pyarmor
```

### "Permission denied"
Run the build script with appropriate permissions or check folder access.

### "Import errors in protected version"
Ensure all dependencies from `requirements.txt` are installed in the target environment.

---

## Technical Details

### PyArmor Protection:
- Uses AES-256 encryption for code
- Runtime decryption with custom loader
- Anti-debugging features
- Restricted mode prevents import from plain scripts

### Bytecode Protection:
- Python's built-in `py_compile`
- Removes source `.py` files
- Wrapper script handles module loading
- Basic obfuscation

---

## Support

For questions about the protection setup, refer to:
- PyArmor documentation: https://pyarmor.readthedocs.io/
- Python compileall: https://docs.python.org/3/library/compileall.html
