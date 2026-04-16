@echo off
title Agente Codificador

echo ============================================
echo   Agente Codificador
echo ============================================
echo.

:: Vai para a pasta onde este BAT esta salvo
cd /d "%~dp0"

:: Detecta qual comando Python esta disponivel
set PYTHON=
py --version >nul 2>&1
if not errorlevel 1 set PYTHON=py

if "%PYTHON%"=="" (
    python --version >nul 2>&1
    if not errorlevel 1 set PYTHON=python
)

if "%PYTHON%"=="" (
    echo [ERRO] Python nao encontrado!
    echo Instale em: https://www.python.org/downloads/
    echo Marque "Add Python to PATH" durante a instalacao.
    pause
    exit /b 1
)
echo [OK] Python encontrado: %PYTHON%

:: Instala dependencias
echo Instalando/verificando dependencias...
%PYTHON% -m pip install -r requirements.txt --quiet
if errorlevel 1 (
    echo [ERRO] Falha ao instalar dependencias.
    pause
    exit /b 1
)
echo [OK] Dependencias prontas.

:: Roda o app
echo.
echo Iniciando o app...
echo.
%PYTHON% app.py

if errorlevel 1 (
    echo.
    echo [AVISO] O programa encerrou com erro.
    pause
)
