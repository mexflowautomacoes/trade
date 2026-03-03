@echo off
echo ============================================
echo   FlowTrader Pro v3 - Atualizacao
echo ============================================
echo.

REM Save current directory
set "APP_DIR=%~dp0"
set "REPO_URL=https://github.com/mexflowautomacoes/trade/archive/refs/heads/main.zip"
set "ZIP_FILE=%APP_DIR%_update.zip"
set "EXTRACT_DIR=%APP_DIR%_update"

echo   Baixando ultima versao do GitHub...
echo.
powershell -Command "try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri '%REPO_URL%' -OutFile '%ZIP_FILE%' -UseBasicParsing } catch { Write-Host '  [ERRO] Falha no download:' $_.Exception.Message; exit 1 }"

if %ERRORLEVEL% neq 0 (
    echo.
    echo   [ERRO] Nao foi possivel baixar a atualizacao.
    echo   Verifique sua conexao com a internet.
    pause
    exit /b 1
)

echo   Download concluido. Extraindo arquivos...
echo.

REM Clean previous extract if exists
if exist "%EXTRACT_DIR%" rmdir /s /q "%EXTRACT_DIR%"

powershell -Command "try { Expand-Archive -Path '%ZIP_FILE%' -DestinationPath '%EXTRACT_DIR%' -Force } catch { Write-Host '  [ERRO] Falha ao extrair:' $_.Exception.Message; exit 1 }"

if %ERRORLEVEL% neq 0 (
    echo   [ERRO] Falha ao extrair o arquivo.
    del "%ZIP_FILE%" >nul 2>&1
    pause
    exit /b 1
)

REM The zip extracts to a subfolder named trade-main
set "SRC_DIR=%EXTRACT_DIR%\trade-main"

if not exist "%SRC_DIR%\flowtrader_server.py" (
    echo   [ERRO] Arquivo flowtrader_server.py nao encontrado no download.
    echo   O repositorio pode ter mudado de estrutura.
    rmdir /s /q "%EXTRACT_DIR%" >nul 2>&1
    del "%ZIP_FILE%" >nul 2>&1
    pause
    exit /b 1
)

echo   Atualizando arquivos...
echo.

REM Update only application files (preserve config.json, .db, etc.)
copy /y "%SRC_DIR%\flowtrader_server.py" "%APP_DIR%flowtrader_server.py" >nul
echo   [OK] flowtrader_server.py
copy /y "%SRC_DIR%\requirements.txt" "%APP_DIR%requirements.txt" >nul
echo   [OK] requirements.txt
copy /y "%SRC_DIR%\setup.bat" "%APP_DIR%setup.bat" >nul
echo   [OK] setup.bat
copy /y "%SRC_DIR%\iniciar_flowtrader.bat" "%APP_DIR%iniciar_flowtrader.bat" >nul
echo   [OK] iniciar_flowtrader.bat
copy /y "%SRC_DIR%\atualizar.bat" "%APP_DIR%atualizar.bat" >nul
echo   [OK] atualizar.bat

echo.
echo   Instalando dependencias (caso haja novas)...
pip install -r "%APP_DIR%requirements.txt" -q >nul 2>&1

echo.
echo   Regenerando dashboard...
python "%APP_DIR%flowtrader_server.py" --generate-html

if %ERRORLEVEL% neq 0 (
    echo   [AVISO] Nao foi possivel regenerar o HTML.
    echo   O dashboard sera regenerado na proxima execucao.
)

REM Cleanup
echo.
echo   Limpando arquivos temporarios...
rmdir /s /q "%EXTRACT_DIR%" >nul 2>&1
del "%ZIP_FILE%" >nul 2>&1

echo.
echo   ============================================
echo   [OK] Atualizacao concluida com sucesso!
echo   ============================================
echo.
echo   Arquivos preservados (nao sobrescritos):
echo     - config.json (suas configuracoes)
echo     - flowtrader_trades.db (historico)
echo.
echo   Reinicie o FlowTrader para usar a nova versao.
echo.
pause
