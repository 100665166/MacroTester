@echo off
rem NOTE: Supports multiple file input

for %%a in (%*) do (
    cscript /nologo %~dp0MacroTester_v1.vbs "%%a"
)

rem For testing only; remove for release
pause