@echo off
chcp 65001 > nul
setlocal enabledelayedexpansion

echo ============================================================
echo  EXTRATOR DE ODDS - INICIALIZADOR
echo ============================================================
echo.

:: Verificar se Python está instalado
echo [1/5] Verificando Python...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ❌ Python não encontrado!
    echo.
    echo 📥 Baixando e instalando Python...
    echo.
    
    :: Baixar instalador do Python
    powershell -Command "& {[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.11.8/python-3.11.8-amd64.exe' -OutFile 'python_installer.exe'}"
    
    if exist python_installer.exe (
        echo ⚙️  Instalando Python (aguarde)...
        start /wait python_installer.exe /quiet InstallAllUsers=0 PrependPath=1 Include_test=0
        del python_installer.exe
        
        echo ✅ Python instalado!
        echo 🔄 Reinicie este script para continuar
        pause
        exit /b 0
    ) else (
        echo ❌ Falha ao baixar Python
        echo.
        echo 💡 Instale manualmente de: https://www.python.org/downloads/
        pause
        exit /b 1
    )
) else (
    for /f "tokens=2" %%i in ('python --version 2^>^&1') do set PYTHON_VER=%%i
    echo ✅ Python !PYTHON_VER! encontrado
)

echo.
echo [2/5] Verificando ambiente virtual...

:: Verificar se venv existe
if not exist "venv\" (
    echo 📦 Criando ambiente virtual...
    python -m venv venv
    if %errorlevel% neq 0 (
        echo ❌ Erro ao criar venv
        pause
        exit /b 1
    )
    echo ✅ Ambiente virtual criado
) else (
    echo ✅ Ambiente virtual já existe
)

echo.
echo [3/5] Ativando ambiente virtual...
call venv\Scripts\activate.bat
if %errorlevel% neq 0 (
    echo ❌ Erro ao ativar venv
    pause
    exit /b 1
)
echo ✅ Ambiente ativado

echo.
echo [4/5] Instalando/Verificando dependências...

:: Instalar dependências
echo 📦 Instalando dependências...
python -m pip install --upgrade pip -q
pip install beautifulsoup4 openpyxl selenium webdriver-manager -q

if %errorlevel% neq 0 (
    echo ⚠️  Algumas dependências podem ter falhado, mas continuando...
)
echo ✅ Dependências prontas

echo.
echo [5/5] Iniciando extrator de odds...
echo.
echo ============================================================
echo.

:: Executar o script Python
python extract_odds.py

echo.
echo ============================================================
echo  Processo finalizado
echo ============================================================
echo.

:: Manter janela aberta
pause
