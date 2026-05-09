@echo off
setlocal

set CSC=
if exist "%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\csc.exe" (
    set CSC=%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\csc.exe
) else if exist "%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\csc.exe" (
    set CSC=%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\csc.exe
) else (
    echo ERROR: csc.exe not found. .NET Framework 4.8 required.
    exit /b 1
)

echo Building VDJSync.exe...
"%CSC%" /target:winexe /out:VDJSync.exe /optimize ^
    /reference:System.dll ^
    /reference:System.Windows.Forms.dll ^
    /reference:System.Drawing.dll ^
    /reference:System.Xml.dll ^
    /reference:System.Net.Http.dll ^
    /reference:System.IO.Compression.dll ^
    /reference:System.IO.Compression.FileSystem.dll ^
    /reference:System.Security.dll ^
    VDJSync.cs

if %ERRORLEVEL% EQU 0 (
    echo Build successful: VDJSync.exe
) else (
    echo Build FAILED.
)

pause
