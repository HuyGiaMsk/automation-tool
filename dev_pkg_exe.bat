@echo off

IF NOT EXIST venv (
    python -m venv venv
    echo Virtual environment created: venv
)

call venv\Scripts\activate.bat

pip install -r requirements.txt

pyinstaller automation_tool.spec

IF EXIST dist (
    xcopy input dist\input /E /H /C /I
    xcopy script dist\script /E /H /C /I
)

deactivate

"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" automation_tool.iss
pause
