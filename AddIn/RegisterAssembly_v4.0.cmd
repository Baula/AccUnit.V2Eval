@echo off

set DotNetToolsPath=%systemroot%\Microsoft.NET\Framework\v4.0.30319

%DotNetToolsPath%\RegAsm.exe %~dp0bin\Debug\AccessCodeLib.AccUnit.V2Eval.AddIn.dll /codebase /tlb:AccessCodeLib.AccUnit.V2Eval.AddIn.tlb

pause