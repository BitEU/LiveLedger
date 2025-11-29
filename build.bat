@echo off
REM BUILD SCRIPT - Save as build.bat
echo Building Windows Terminal Spreadsheet Calculator...

REM Set MSVC paths directly to avoid PATH length issues in cmd.exe
set "VSBASE=C:\Program Files (x86)\Microsoft Visual Studio\2019\BuildTools"
set "VCTOOLS=%VSBASE%\VC\Tools\MSVC\14.29.30133\bin\Hostx64\x64"
set "WINSDK=C:\Program Files (x86)\Windows Kits\10\bin\10.0.19041.0\x64"
set "INCLUDE=%VSBASE%\VC\Tools\MSVC\14.29.30133\include;C:\Program Files (x86)\Windows Kits\10\Include\10.0.19041.0\ucrt;C:\Program Files (x86)\Windows Kits\10\Include\10.0.19041.0\um;C:\Program Files (x86)\Windows Kits\10\Include\10.0.19041.0\shared"
set "LIB=%VSBASE%\VC\Tools\MSVC\14.29.30133\lib\x64;C:\Program Files (x86)\Windows Kits\10\Lib\10.0.19041.0\ucrt\x64;C:\Program Files (x86)\Windows Kits\10\Lib\10.0.19041.0\um\x64"

REM Check if compiler exists
if exist "%VCTOOLS%\cl.exe" (
    echo Using MSVC compiler...
    echo Compiling resource file...
    "%WINSDK%\rc.exe" resource.rc
    if %ERRORLEVEL% EQU 0 (
        echo Compiling and linking with icon...
        "%VCTOOLS%\cl.exe" /O2 /W3 /TC main.c sheet.c console.c charts.c /Fe:LL.exe /link resource.res user32.lib
    ) else (
        echo Warning: Resource compilation failed, building without icon...
        "%VCTOOLS%\cl.exe" /O2 /W3 /TC main.c sheet.c console.c charts.c /Fe:LL.exe /link user32.lib
    )
) else (
    echo Error: Visual Studio compiler not found!
    echo Please install Visual Studio Build Tools or adjust paths in build.bat
    exit /b 1
)

if %ERRORLEVEL% EQU 0 (
    echo Build successful!
    echo Run LL.exe to start the spreadsheet
    REM Clean up temporary resource files
    if exist resource.res del resource.res >nul 2>nul
    if exist sheet.obj del sheet.obj >nul 2>nul
    if exist main.obj del main.obj >nul 2>nul
    if exist console.obj del console.obj >nul 2>nul
    if exist charts.obj del charts.obj >nul 2>nul
) else (
    echo Build failed!
    exit /b 1
)