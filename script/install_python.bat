@echo off

set DEFAULT_INSTALL_PATH=%USERPROFILE%\AppData\Local\Programs\Python

:install_python
    echo --------------------------------------------------------------------------------------------------------------
    echo Start trying to install Python
    set PYTHON_VERSION=3.10.10

    call python --version >NUL 2>&1
    if %errorlevel%==0 (
        echo Python is already installed at:
        call where python.exe
    ) else (
        mkdir %DEFAULT_INSTALL_PATH%
        echo Will install Python at %DEFAULT_INSTALL_PATH%
        curl -o python-installer.exe https://www.python.org/ftp/python/%PYTHON_VERSION%/python-%PYTHON_VERSION%-amd64.exe
        python-installer.exe /quiet InstallAllUsers=1 PrependPath=1 TargetDir=%DEFAULT_INSTALL_PATH%

        set python_script_path=%DEFAULT_INSTALL_PATH%\Python310\Scripts\
        set python_path=%DEFAULT_INSTALL_PATH%\Python310\
        setx PATH "%python_path%;%python_script_path%;%PATH%" /M
        echo Python has been added to the system PATH environment variable.
        del python-installer.exe
    )
    call python -m pip install --upgrade pip
    call pip install certifi
    call pip install requests==2.28.2

:eof