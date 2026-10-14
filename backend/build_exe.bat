@echo off
setlocal
cd /d "%~dp0"
python -m pip install -r requirements-build.txt
if errorlevel 1 goto failed
python -m PyInstaller --noconfirm --onefile --windowed --name ConciliacaoDMA --icon "..\frontend\foco.ico" --paths . --add-data "..\frontend\foco-logo.png;frontend" --add-data "..\frontend\foco.ico;frontend" --distpath "..\dist" --workpath "..\build\pyinstaller-rebuild" desktop\main.py
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
