@echo off

setlocal enabledelayedexpansion

:find_onedrive_mapping_folder
set "searchTerm=OneDrive "
set "OneDriveFolder=Documents"
for /d %%D in ("%USERPROFILE%\*") do (
    for %%F in ("%%~nxD") do (
        set "folderName=%%~nF"
        if /I "!folderName:~0,9!"=="!searchTerm!" (
            set "OneDriveFolder=%%~nF"
            goto :set_main_paths
        )
    )
)

:set_main_paths
set SOURCE_PATH=%cd%\automation_tool

:download_source_code
    echo --------------------------------------------------------------------------------------------------------------
    echo Start trying to install the source code at %SOURCE_PATH%

    if exist "%SOURCE_PATH%" (
        echo Already contains source at %SOURCE_PATH%
        goto :eof
    ) else (
        echo Start dumping temp download source .py
        (
            echo B^=print
            echo import os as A^,zipfile as H^,requests as I
            echo E^=A.path.dirname^(A.path.abspath^(__file__^)^)
            echo F^='automation_tool'
            echo J^=A.path.join^(E^,F^)
            echo def C^(^):
            echo 	if A.path.exists^(J^):B^('Already containing the source code'^)^;return
            echo 	K^=f^"https://github.com/HuyGiaMsk/automation-tool/archive/main.zip^"^;B^('Start download source'^)^;G^=I.get^(url^=K^,verify^=False^)
            echo 	if G.status_code^=^=200:
            echo 		C^=E^;D^=A.path.join^(C^,'automation-tool-main.zip'^)
            echo 		with open^(D^,'wb'^)as L:L.write^(G.content^)
            echo 		B^('Download source successfully'^)
            echo 		with H.ZipFile^(D^,'r'^)as M:M.extractall^(C^)
            echo 		A.rename^(A.path.join^(C^,'automation-tool-main'^)^,A.path.join^(C^,F^)^)^;A.remove^(D^)^;B^(f^"Extracted source code and placed it in ^{C^}^"^)
            echo 	else:B^('Failed to download the source'^)
            echo if __name__^=^='__main__':C^(^)
        ) > temporary_download_source.py

        echo Invoke temp download source .py
        call python temporary_download_source.py

        echo Invoke temp download source .py complete
        del temporary_download_source.py
    )

endlocal
:eof