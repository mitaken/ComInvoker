@echo off
"%ProgramFiles(x86)%\Microsoft Visual Studio\2017\Professional\MSBuild\15.0\Bin\MSBuild.exe" %~dp0\ComInvoker.sln /p:Configuration=Release /t:Clean;Build
pause
