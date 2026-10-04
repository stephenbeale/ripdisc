@echo off
REM Build script for RipDisc C# application

echo Building RipDisc...
cd /d "%~dp0"
dotnet build RipDisc.sln -c Release

if %ERRORLEVEL% EQU 0 (
    echo.
    echo Build successful!
    echo Executable location: RipDisc.Cli\bin\Release\net8.0-windows\RipDisc.exe
    echo Run the tests with: dotnet test RipDisc.sln
    echo.
) else (
    echo.
    echo Build failed!
    echo.
    exit /b 1
)
