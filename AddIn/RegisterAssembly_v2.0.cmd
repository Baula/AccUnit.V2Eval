@echo off

set DotNetToolsPath=%systemroot%\Microsoft.NET\Framework\v2.0.50727

%DotNetToolsPath%\RegAsm.exe %~dp0bin\Debug\AccessCodeLib.AccUnit.V2Eval.AddIn.dll /codebase /tlb:AccessCodeLib.AccUnit.V2Eval.AddIn.tlb

pause