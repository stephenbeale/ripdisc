@echo off
REM Build script for RipDisc C# application

echo Building RipDisc...
cd RipDisc
dotnet build -c Release

if %ERRORLEVEL% EQU 0 (
    echo.
    echo Build successful!
    echo Executable location: RipDisc\bin\Release\net8.0-windows\RipDisc.exe
    echo.
) else (
    echo.
    echo Build failed!
    echo.
    exit /b 1
)
