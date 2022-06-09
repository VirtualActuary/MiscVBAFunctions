@echo off

:: TODO: add zebra-vba-packager in requirements.txt
::pip3 install -r "%~dp0\requirements.txt" ' 
python "%~dp0\compile.py"

pip3 install -r "%~dp0\..\docs\requirements.txt"

pushd "%~dp0\..\docs"
    mkdocs build -f mkdocs.yml
popd

if /i "%comspec% /c ``%~0` `" equ "%cmdcmdline:"=`%" pause