@echo off
echo ============================================
echo   FlowTrader Pro v3.0
echo ============================================
echo.

REM Check Python
python --version >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo   [ERRO] Python nao encontrado!
    echo   Execute setup.bat primeiro.
    pause
    exit /b 1
)

REM Check dependencies
python -c "import xlwings; import websockets" >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo   [AVISO] Dependencias nao instaladas.
    echo   Executando setup.bat...
    echo.
    call "%~dp0setup.bat"
    if %ERRORLEVEL% neq 0 exit /b 1
    echo.
)

echo   Dica: abra o ProfitChart e o Excel
echo   antes ou durante a execucao.
echo.

REM Start server (usa config.json se existir, senao auto-detecta workbook aberto)
REM Use Configuracoes no dashboard para definir workbook, aba, etc.
python "%~dp0flowtrader_server.py"

echo.
echo   Servidor parou. Pressione qualquer tecla para fechar.
pause >nul
