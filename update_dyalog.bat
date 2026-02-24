@echo off
if "%~1"=="" (set VER=20.0) else (set VER=%~1)
set SRC=C:\Program Files\Dyalog\Dyalog APL-64 %VER% Unicode
set DST=%~dp0Dyalog

echo Updating Dyalog folder from "%SRC%"...

for %%F in (
    bridge200-64_unicode.dll
    dyalog.exe
    Dyalog.Net.Bridge.dll
    dyalognet.dll
) do (
    echo   %%F
    copy /Y "%SRC%\%%F" "%DST%\%%F" >nul
)

echo Done.
