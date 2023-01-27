@echo off
python "%~dp0%~n0.py"
if /i "%comspec% /c %~0 " equ "%cmdcmdline:"=%" pause
