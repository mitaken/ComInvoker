@echo off
"%ProgramFiles(x86)%\MSBuild\14.0\Bin\msbuild.exe" %~dp0\ComInvoker.sln /p:Configuration=Release /t:Clean;Build
pause
