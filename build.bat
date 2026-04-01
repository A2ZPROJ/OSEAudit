@echo off
title Compilador - OSEAudit
set ISCC="C:\Program Files\Inno Setup 7\ISCC.exe"
set GH="C:\Program Files\GitHub CLI\gh.exe"

echo.
echo === COMPILADOR - OSEAudit (A2Z Projetos) ===
echo.

:: Localiza o Python instalado
set PYTHON=C:\Users\%USERNAME%\AppData\Local\Programs\Python\Python312\python.exe
if not exist "%PYTHON%" (
    echo [ERRO] Python 3.12 nao encontrado em %PYTHON%
    echo        Instale o Python 3.10+ em https://www.python.org/downloads/
    echo        e marque "Add Python to PATH" durante a instalacao.
    pause
    exit /b 1
)

echo [OK] Python encontrado:
"%PYTHON%" --version
echo.

:: Vai para a pasta do script
cd /d "%~dp0"

:: Instala dependencias
echo [1/5] Instalando dependencias...
"%PYTHON%" -m pip install --upgrade pip --quiet
"%PYTHON%" -m pip install pdfplumber openpyxl pyinstaller fpdf2 --quiet

if errorlevel 1 (
    echo [ERRO] Falha ao instalar dependencias.
    pause
    exit /b 1
)
echo [OK] Dependencias instaladas.
echo.

:: Compila o executavel
echo [2/5] Compilando executavel...
echo       Aguarde, isso pode demorar alguns minutos...
echo.

"%PYTHON%" -m PyInstaller --onedir --windowed --name "OSEAudit" --icon "assets\oseaudit.ico" --add-data "comparador_ose.py;." --add-data "splash.py;." --add-data "assets;assets" --hidden-import pdfplumber --hidden-import pdfminer --hidden-import pdfminer.high_level --hidden-import pdfminer.layout --hidden-import openpyxl --hidden-import openpyxl.styles --hidden-import openpyxl.utils --hidden-import PIL --hidden-import fpdf --collect-all pdfplumber --collect-all pdfminer interface.py

if errorlevel 1 (
    echo.
    echo [ERRO] Falha na compilacao. Verifique as mensagens acima.
    pause
    exit /b 1
)
echo [OK] Compilacao concluida.
echo.

:: Gera o instalador com Inno Setup
echo [3/5] Gerando instalador...
%ISCC% "OSEAudit.iss"

if errorlevel 1 (
    echo [ERRO] Falha ao gerar instalador.
    pause
    exit /b 1
)
echo [OK] Instalador gerado: OSEAudit_Setup.exe
echo.

:: Limpa arquivos temporarios
echo [4/5] Limpando arquivos temporarios...
rmdir /s /q build >nul 2>&1
rmdir /s /q dist  >nul 2>&1
del /q "OSEAudit.spec" >nul 2>&1
echo [OK] Limpeza concluida.
echo.

:: Publica release no GitHub (requer GitHub CLI instalado e autenticado)
echo [5/5] Publicando release no GitHub...
if not exist %GH% (
    echo [AVISO] GitHub CLI nao encontrado. Pulando publicacao automatica.
    echo         Baixe em: https://cli.github.com
    goto :fim
)

:: Le a versao do interface.py
for /f "tokens=3 delims= " %%v in ('findstr "VERSAO" interface.py') do (
    set RAW_VERSAO=%%v
    goto :versao_ok
)
:versao_ok
set VERSAO=%RAW_VERSAO:"=%

echo Versao detectada: %VERSAO%
%GH% release create "v%VERSAO%" "OSEAudit_Setup.exe" --repo A2ZPROJ/OSEAudit --title "OSEAudit v%VERSAO%" --notes "Release v%VERSAO%"

if errorlevel 1 (
    echo [AVISO] Falha ao publicar no GitHub. Verifique se esta autenticado com: gh auth login
) else (
    echo [OK] Release v%VERSAO% publicada no GitHub.
)

:fim
echo.
echo === Tudo pronto! Arquivo gerado: OSEAudit_Setup.exe ===
echo.
echo Pressione qualquer tecla para abrir o instalador...
pause >nul
start "" "OSEAudit_Setup.exe"
