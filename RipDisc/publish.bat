@echo off
REM Publish script for RipDisc C# application
REM Creates a self-contained executable

echo Publishing RipDisc...
cd RipDisc
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true

if %ERRORLEVEL% EQU 0 (
    echo.
    echo Publish successful!
    echo Self-contained executable: RipDisc\bin\Release\net8.0-windows\win-x64\publish\RipDisc.exe
    echo.
    echo You can copy RipDisc.exe to any location and run it without requiring .NET installation.
    echo.
) else (
    echo.
    echo Publish failed!
    echo.
    exit /b 1
)
