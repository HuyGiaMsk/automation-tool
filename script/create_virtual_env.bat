@echo off

:set_main_paths
set SOURCE_PATH=%cd%
set VENV_PATH=%SOURCE_PATH%\venv

:create_virtual_env
    echo --------------------------------------------------------------------------------------------------------------
    echo Start trying to install virtual environments
    echo %VENV_PATH%
    if exist "%VENV_PATH%" (
        echo Virtual environment is already installed in %VENV_PATH%
        goto :eof
    ) else (
        call python -m venv "%VENV_PATH%"
        call "%VENV_PATH%"\Scripts\activate

        echo Install defined packages
        pip install -r "%SOURCE_PATH%\requirements.txt"

        call "%VENV_PATH%"\Scripts\deactivate
        echo Install virtual env successfully at %VENV_PATH%
    )
:eof