@echo off

python "%~dp0\scripts\compile.py"

if /i "%comspec% /c ``%~0` `" equ "%cmdcmdline:"=`%" pause