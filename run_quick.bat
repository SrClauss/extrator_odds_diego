@echo off
chcp 65001 > nul
echo ============================================================
echo  EXTRATOR DE ODDS
echo ============================================================
echo.

:: Ativar venv
if exist "venv\Scripts\activate.bat" (
    call venv\Scripts\activate.bat
    echo ✅ Ambiente virtual ativado
    echo.
    
    :: Executar script
    python extract_odds.py
    
    echo.
    pause
) else (
    echo ❌ Ambiente virtual não encontrado!
    echo.
    echo 💡 Execute primeiro: run_extract_odds.bat
    echo.
    pause
)
