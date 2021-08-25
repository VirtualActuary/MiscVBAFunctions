@echo off

python "%~dp0\scripts\decompile.py"

if /i "%comspec% /c ``%~0` `" equ "%cmdcmdline:"=`%" pause