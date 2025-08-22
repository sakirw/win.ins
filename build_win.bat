
@echo off
setlocal ENABLEDELAYEDEXPANSION
if not exist .venv ( py -3 -m venv .venv )
call .venv\Scripts\activate
python -m pip install --upgrade pip
pip install -r requirements.txt
if exist icon.png ( python convert_icon.py )
rmdir /s /q build 2>nul
rmdir /s /q dist 2>nul
del SatisRaporGuncelleme.spec 2>nul
set ICON_ARG=
if exist icon.ico ( set ICON_ARG=--icon icon.ico )
pyinstaller --noconfirm --windowed --name "SatisRaporGuncelleme" %ICON_ARG% ^
  satis_rapor_guncelleme.py
echo Build tamamlandi. dist\SatisRaporGuncelleme altinda.
