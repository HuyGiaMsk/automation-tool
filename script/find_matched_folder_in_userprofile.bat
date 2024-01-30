@echo off
setlocal enabledelayedexpansion

:find_matched_folder_in_userprofile
set "searchTerm=%~1"
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
echo OneDrive folder found: %OneDriveFolder%
goto :eof

REM call :find_matched_folder_in_userprofile "OneDrive"

