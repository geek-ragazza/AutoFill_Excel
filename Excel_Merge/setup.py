# -*- coding: utf-8 -*-

import sys
from cx_Freeze import setup, Executable

# Dependencies are automatically detected, but it might need fine tuning.
build_exe_options = {
    "includes": ["Tkinter"],
    "append_script_to_exe":False,
    "compressed":True,
    "path": sys.path,
    "copy_dependent_files":True,
    "create_shared_zip":True,
    "include_in_shared_zip":True,
    "optimize":2,
    "build_exe":"dist/bin",
    }

# GUI applications require a different base on Windows (the default is for a
# console application).
base = None
if sys.platform == "win32":
    base = "Win32GUI"

setup(
    name = "simple_Tkinter",
    version = "0.1",
    description = "Sample cx_Freeze Tkinter script",
    options = {"build_exe": build_exe_options},
    executables = [Executable(
        script = "Panel_Merge_console.py",
        base = base,
        targetDir = "dist",
        initScript = None,
        targetName = "foo.exe",
        compress = True,
        copyDependentFiles = True,
        appendScriptToExe = True,
        appendScriptToLibrary = True,
        )])

