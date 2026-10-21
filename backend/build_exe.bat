@echo off
setlocal
cd /d "%~dp0"
python -m pip install -r requirements-build.txt
if errorlevel 1 goto failed
python -m PyInstaller --noconfirm --distpath "..\dist" --workpath "..\build\pyinstaller-rebuild" ConciliacaoDMA.spec
if errorlevel 1 goto failed
echo.
echo Executavel criado em: dist\ConciliacaoDMA.exe
pause
exit /b 0
:failed
echo.
echo Falha ao instalar dependencias ou criar o executavel.
pause
exit /b 1
