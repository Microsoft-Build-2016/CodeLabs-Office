@echo off
%~d0
cd "%~dp0"

echo.
echo ===================================================
echo Install Visual Studio Code Snippets for the module
echo ===================================================
echo.
"c:\Program Files (x86)\Microsoft Visual Studio 14.0\Common7\IDE\VSIXInstaller.exe" /q /a "%~dp0snippets\OfficeSnippets.vsix"
