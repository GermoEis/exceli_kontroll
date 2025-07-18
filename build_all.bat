@echo off
echo Alustan EXE-de loomist...

python -m PyInstaller Exceli_kontroll.spec
python -m PyInstaller excelite_vordlus.spec
python -m PyInstaller Kontroll.spec
python -m PyInstaller kuupaeva_kontroll.spec
python -m PyInstaller xml_muutmine.spec

echo Valmis! Exe failid on kaustas dist\
pause
