# Script de Build Local (Windows)

@echo off
chcp 65001 > nul
echo ============================================================
echo  BUILD DO EXTRATOR DE ODDS
echo ============================================================
echo.

:: Ativar venv se existir
if exist "venv\Scripts\activate.bat" (
    call venv\Scripts\activate.bat
)

echo [1/4] Instalando ferramentas de build...
pip install pyinstaller pyarmor -q

echo.
echo [2/4] Ofuscando código com PyArmor...
pyarmor gen -O dist-obf extract_odds.py
if %errorlevel% neq 0 (
    echo ⚠️  PyArmor falhou, usando código original
    set USE_ORIGINAL=1
) else (
    echo ✅ Código ofuscado
    set USE_ORIGINAL=0
)

echo.
echo [3/4] Compilando com PyInstaller...

if "%USE_ORIGINAL%"=="1" (
    pyinstaller --onefile --name "ExtractOdds" ^
        --add-data "README.md;." ^
        --hidden-import=selenium ^
        --hidden-import=webdriver_manager ^
        --hidden-import=bs4 ^
        --hidden-import=openpyxl ^
        extract_odds.py
) else (
    pyinstaller --onefile --name "ExtractOdds" ^
        --add-data "README.md;." ^
        --hidden-import=selenium ^
        --hidden-import=webdriver_manager ^
        --hidden-import=bs4 ^
        --hidden-import=openpyxl ^
        dist-obf\extract_odds.py
)

if %errorlevel% neq 0 (
    echo ❌ Erro no build
    pause
    exit /b 1
)

echo.
echo [4/4] Copiando arquivos auxiliares...
copy run_extract_odds.bat dist\
copy run_quick.bat dist\
copy requirements_minimo.txt dist\requirements.txt
copy README.md dist\

echo.
echo ============================================================
echo ✅ BUILD CONCLUÍDO
echo ============================================================
echo.
echo 📁 Arquivos em: dist\
echo 📦 Executável: dist\ExtractOdds.exe
echo.
pause
