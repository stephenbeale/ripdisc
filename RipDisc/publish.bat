@echo off
REM Publish script for RipDisc C# application
REM Creates a self-contained executable

echo Publishing RipDisc...
cd /d "%~dp0"
dotnet publish RipDisc.Cli\RipDisc.Cli.csproj -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true

if %ERRORLEVEL% EQU 0 (
    echo.
    echo Publish successful!
    echo Self-contained executable: RipDisc.Cli\bin\Release\net8.0-windows\win-x64\publish\RipDisc.exe
    echo.
    echo You can copy RipDisc.exe to any location and run it without requiring .NET installation.
    echo Put ripdisc-config.json next to it ^(or in a parent folder^) to set tool paths and drives.
    echo.
) else (
    echo.
    echo Publish failed!
    echo.
    exit /b 1
)
