@echo off
setlocal

rem Builds the DUnitX suite for Win32 and Win64 and runs it on both.
rem From a RAD Studio command prompt, where BDS is set, it uses that Delphi;
rem otherwise it uses the newest RAD Studio in the registry.

set "ROOT=%~dp0"
set "TEST_PROJECT=%ROOT%Tests\Source\Office4D.Tests.dproj"

if defined BDS (
    set "BDS_ROOT=%BDS%\"
) else (
    call :FindNewestDelphi
)
if not defined BDS_ROOT (
    echo [build] No RAD Studio found in the registry.
    exit /b 1
)
if not exist "%BDS_ROOT%bin\rsvars.bat" (
    echo [build] rsvars.bat not found in "%BDS_ROOT%bin".
    exit /b 1
)
call "%BDS_ROOT%bin\rsvars.bat"
echo [build] Using "%BDS%"

for %%P in (Win32 Win64) do (
    echo.
    echo === Building tests ^(%%P^) ===
    msbuild "%TEST_PROJECT%" /t:Build /p:Config=Debug /p:Platform=%%P /v:m
    if errorlevel 1 (
        echo === BUILD FAILED ^(%%P^) ===
        exit /b 1
    )
)

for %%P in (Win32 Win64) do (
    echo.
    echo === Running tests ^(%%P^) ===
    pushd "%ROOT%Tests\Source\%%P\Debug"
    ".\Office4D.Tests.exe" --exitbehavior:Continue
    if errorlevel 1 (
        popd
        echo === TESTS FAILED ^(%%P^) ===
        exit /b 1
    )
    popd
)

echo.
echo === All tests passed on Win32 and Win64 ===
exit /b 0

rem Picks the installed version with the highest major number, so 37.0 wins
rem over 23.0 whatever order the registry lists them in.
:FindNewestDelphi
set BEST_MAJOR=0
for /f "delims=" %%K in ('reg query "HKLM\SOFTWARE\WOW6432Node\Embarcadero\BDS" 2^>nul ^| findstr /r "\\[0-9][0-9]*\.[0-9]*$"') do call :ConsiderVersion "%%K"
if not defined BDS_KEY exit /b 0
for /f "tokens=2,*" %%A in ('reg query "%BDS_KEY%" /v RootDir 2^>nul ^| findstr /i "RootDir"') do set "BDS_ROOT=%%B"
exit /b 0

:ConsiderVersion
for /f "delims=." %%M in ("%~nx1") do set /a MAJOR=%%M
if %MAJOR% gtr %BEST_MAJOR% (
    set BEST_MAJOR=%MAJOR%
    set "BDS_KEY=%~1"
)
exit /b 0
