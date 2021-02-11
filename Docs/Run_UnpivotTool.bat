@echo off

Set ProgramPath=UnpivotTool.exe

If Exist ..\UnpivotTool.exe     Set ProgramPath=..\UnpivotTool.exe
If Exist ..\Bin\UnpivotTool.exe Set ProgramPath=..\Bin\UnpivotTool.exe

Rem The following will process file ExamplePivotTable.txt, creating file ExamplePivotTable_Unpivot.txt

@echo on
%ProgramPath% /I:ExamplePivotTable.txt /F:2 /B /N