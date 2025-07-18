@echo off
echo Alustan EXE-de loomist...

python -m PyInstaller Exceli_kontroll.spec
python -m PyInstaller excelite_võrdlus.spec
python -m PyInstaller Kontroll.spec
python -m PyInstaller kuupäeva_kontroll.spec
python -m PyInstaller xml_muutmine.spec

echo Valmis! Exe failid on kaustas dist\
pause
