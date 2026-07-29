@echo off
setlocal
cd /d "%~dp0"

set "WGA_PYTHON=%~dp0.runtime\python-standalone\python\python.exe"

if not exist "%WGA_PYTHON%" (
    echo [WGA TEST] Local Python runtime was not found:
    echo %WGA_PYTHON%
    echo.
    pause
    exit /b 1
)

echo [WGA TEST] Starting Windows 10 Optimization screen...
"%WGA_PYTHON%" -c "import tkinter as tk; from optimization_ui import OptimizationUI; root=tk.Tk(); OptimizationUI(root, root.destroy); root.mainloop()"

if errorlevel 1 (
    echo.
    echo [WGA TEST] The Optimization screen stopped with an error.
    pause
    exit /b 1
)

endlocal
