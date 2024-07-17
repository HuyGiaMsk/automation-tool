@echo off
IF EXIST build (
    rmdir /s /q build
)

IF EXIST dist (
    rmdir /s /q dist
)

IF NOT EXIST venv (
    python -m venv venv
    echo Virtual environment created: venv
)

call venv\Scripts\activate.bat

pip install -r requirements.txt

pyinstaller automation_tool.spec

IF EXIST dist (
    xcopy input dist\input /E /H /C /I
)

echo Checking Inno Setup path...
if exist "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" (
    echo Inno Setup executable found.
) else (
    echo Inno Setup executable not found.
    pause
    exit /b 1
)

echo Starting Inno Setup...
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" automation_tool.iss
IF ERRORLEVEL 1 (
    echo Inno Setup failed
    pause
    exit /b 1
)
echo Inno Setup completed.

call venv\Scripts\deactivate.bat
pause
