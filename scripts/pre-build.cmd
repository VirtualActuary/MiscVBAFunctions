@echo off

pip3 install -r "%~dp0\..\docs\requirements.txt"

pushd "%~dp0\..\docs"
    mkdocs build -f mkdocs.yml
popd

python "%~dp0\compile.py"


if /i "%comspec% /c ``%~0` `" equ "%cmdcmdline:"=`%" pause