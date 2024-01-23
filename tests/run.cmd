@echo off
python -m unittest --verbose
if /i "%comspec% /c %~0 " equ "%cmdcmdline:"=%" pause
