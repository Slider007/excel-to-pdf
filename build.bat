@echo off
cd /d "%~dp0"

if exist PdfSharp.dll goto compile
echo Downloading PdfSharp...
powershell -NoProfile -Command "[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; Invoke-WebRequest 'https://www.nuget.org/api/v2/package/PdfSharp/1.50.5147' -OutFile p.zip; Add-Type -AssemblyName System.IO.Compression.FileSystem; $z=[IO.Compression.ZipFile]::OpenRead((Join-Path (Get-Location) 'p.zip')); $e=$z.Entries | Where-Object { $_.FullName -match '^lib/net\d+/PdfSharp\.dll$' } | Select-Object -First 1; [IO.Compression.ZipFileExtensions]::ExtractToFile($e,(Join-Path (Get-Location) 'PdfSharp.dll'),$true); $z.Dispose(); Remove-Item p.zip"
if not exist PdfSharp.dll goto nodll

:compile
set CSC=%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe
"%CSC%" /nologo /codepage:65001 /win32icon:icon.ico /target:winexe /out:"Excel to PDF.exe" /r:PdfSharp.dll /r:Microsoft.CSharp.dll /r:System.Core.dll /r:System.Windows.Forms.dll /r:System.Drawing.dll /resource:PdfSharp.dll,PdfSharp.dll Program.cs
if errorlevel 1 goto fail
echo.
echo DONE: Excel to PDF.exe
pause
exit /b 0

:nodll
echo Could not download PdfSharp.dll. Put it next to build.bat and run again.
pause
exit /b 1

:fail
echo BUILD FAILED - see errors above
pause
exit /b 1
