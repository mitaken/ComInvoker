@echo off
%~dp0\nuget.exe pack .\ComInvoker\ComInvoker.csproj -Symbols -Properties Configuration=Release
pause
