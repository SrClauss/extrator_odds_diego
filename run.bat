@echo off
chcp 65001 >nul
title Projeto Diego

echo ============================================
echo  Atualizando repositorio...
echo ============================================
git fetch --all
git reset --hard origin/master
if %errorlevel% neq 0 (
    echo ERRO: Falha ao atualizar o repositorio.
    pause
    exit /b 1
)
echo Repositorio atualizado com sucesso!
echo.

echo ============================================
echo  Verificando ambiente virtual...
echo ============================================
if not exist "venv\Scripts\activate.bat" (
    echo Criando venv...
    python -m venv venv
    if %errorlevel% neq 0 (
        echo ERRO: Falha ao criar a venv. Verifique se o Python esta instalado.
        pause
        exit /b 1
    )
    echo Venv criada com sucesso!
) else (
    echo Venv ja existe.
)
echo.

echo ============================================
echo  Ativando venv e instalando dependencias...
echo ============================================
call venv\Scripts\activate.bat
if not exist "venv\.installed" (
    echo Instalando dependencias...
    pip install -r requirements.txt
    if %errorlevel% neq 0 (
        echo ERRO: Falha ao instalar as dependencias.
        pause
        exit /b 1
    )
    echo. > venv\.installed
    echo Dependencias instaladas com sucesso!
) else (
    echo Dependencias ja instaladas.
)
echo.

echo ============================================
echo  Executando script...
echo ============================================
python extract_odds.py
echo.

echo ============================================
echo  Script finalizado.
echo ============================================
echo.
set /p FECHAR="Deseja fechar a janela? (s/n): "
if /i "%FECHAR%"=="s" exit
if /i "%FECHAR%"=="sim" exit
cmd /k
