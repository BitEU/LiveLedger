@echo off
REM BUILD SCRIPT FOR TESTS - Save as build_tests.bat
echo Building LiveLedger Unit Tests...

REM Set MSVC paths directly to avoid PATH length issues in cmd.exe
set "VSBASE=C:\Program Files (x86)\Microsoft Visual Studio\2019\BuildTools"
set "VCTOOLS=%VSBASE%\VC\Tools\MSVC\14.29.30133\bin\Hostx64\x64"
set "WINSDK=C:\Program Files (x86)\Windows Kits\10\bin\10.0.19041.0\x64"
set "INCLUDE=%VSBASE%\VC\Tools\MSVC\14.29.30133\include;C:\Program Files (x86)\Windows Kits\10\Include\10.0.19041.0\ucrt;C:\Program Files (x86)\Windows Kits\10\Include\10.0.19041.0\um;C:\Program Files (x86)\Windows Kits\10\Include\10.0.19041.0\shared"
set "LIB=%VSBASE%\VC\Tools\MSVC\14.29.30133\lib\x64;C:\Program Files (x86)\Windows Kits\10\Lib\10.0.19041.0\ucrt\x64;C:\Program Files (x86)\Windows Kits\10\Lib\10.0.19041.0\um\x64"

REM Check if compiler exists
if exist "%VCTOOLS%\cl.exe" (
    echo Using MSVC compiler...
    echo Compiling basic test suite...
    "%VCTOOLS%\cl.exe" /O2 /W3 /TC test_liveledger.c sheet.c console.c charts.c /Fe:test_liveledger.exe /link user32.lib
    
    if %ERRORLEVEL% EQU 0 (
        echo Basic tests build successful!
        REM Clean up temporary object files
        if exist sheet.obj del sheet.obj >nul 2>nul
        if exist test_liveledger.obj del test_liveledger.obj >nul 2>nul
        if exist console.obj del console.obj >nul 2>nul
        if exist charts.obj del charts.obj >nul 2>nul
        
        echo.
        echo Compiling advanced test suite...
        "%VCTOOLS%\cl.exe" /O2 /W3 /TC test_liveledger_advanced.c sheet.c console.c charts.c /Fe:test_liveledger_advanced.exe /link user32.lib
        
        if %ERRORLEVEL% EQU 0 (
            echo Advanced tests build successful!
            echo.
            echo Both test suites built successfully!
            echo Run test_liveledger.exe for basic tests
            echo Run test_liveledger_advanced.exe for advanced tests
            REM Clean up temporary object files
            if exist sheet.obj del sheet.obj >nul 2>nul
            if exist test_liveledger_advanced.obj del test_liveledger_advanced.obj >nul 2>nul
            if exist console.obj del console.obj >nul 2>nul
            if exist charts.obj del charts.obj >nul 2>nul
        ) else (
            echo Advanced tests build failed!
            exit /b 1
        )
    ) else (
        echo Basic tests build failed!
        exit /b 1
    )
) else (
    echo Error: Visual Studio compiler not found!
    echo Please install Visual Studio Build Tools or adjust paths in build_tests.bat
    exit /b 1
)
