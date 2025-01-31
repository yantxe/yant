from cx_Freeze import setup, Executable
import sys
import os

# Get Qt's QML import path
#qt_qml_path = os.path.join(sys.prefix, 'Lib', 'site-packages', 'PyQt5', 'Qt', 'qml')
qt_qml_path = r"E:\robo\stopwatch\myenv\Lib\site-packages\PyQt5\Qt5\qml"
# Include the PyQt5 and other necessary packages
build_options = {
    "packages": ["os", "PyQt5", "pandas", "docx"],
    "excludes": ["tkinter"],
    "include_files": [
        qt_qml_path,  # Manually include the QML files
    ]
}

# Set the base to "Win32GUI" for Windows GUI applications (no terminal)
base = None
if sys.platform == "win32":
    base = "Win32GUI"

# Define the executable and its settings
executables = [
    Executable(
        "stopwatch.py",  # Your Python script
        base=base,
        target_name="stopwatch.exe",  # Name of the executable to be created
        icon=None,  # Specify path to an icon if needed
    )
]

# Setup function to create the executable
setup(
    name="Stopwatch App",
    version="1.0",
    description="A Stopwatch Application",
    options={"build_exe": build_options},
    executables=executables,
)
