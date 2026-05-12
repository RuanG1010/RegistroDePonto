@echo off
setlocal
cd /d "%~dp0"

where py >nul 2>nul
if %errorlevel%==0 (
    set "PY=py -3"
) else (
    set "PY=python"
)

echo Instalando dependencias...
%PY% -m pip install --upgrade pip
if errorlevel 1 goto erro

%PY% -m pip install -r requirements.txt
if errorlevel 1 goto erro

echo.
echo Gerando executavel...
%PY% -m PyInstaller --noconfirm --clean registro_ponto_manual.spec
if errorlevel 1 goto erro

echo.
echo Pronto.
echo Executavel gerado em: dist\RegistroPontoManual\RegistroPontoManual.exe
pause
exit /b 0

:erro
echo.
echo O build falhou. Confira se o Python esta instalado e se este terminal tem permissao para instalar pacotes.
pause
exit /b 1
