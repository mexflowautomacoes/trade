@echo off
echo ============================================
echo   FlowTrader Pro v3 - Setup
echo ============================================
echo.

REM Check Python installation
python --version >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo   [ERRO] Python nao encontrado!
    echo   Instale Python 3.8+ de https://python.org
    echo   Marque "Add Python to PATH" durante a instalacao.
    pause
    exit /b 1
)

echo   [OK] Python encontrado:
python --version
echo.

echo   Instalando dependencias...
echo.
pip install -r "%~dp0requirements.txt"

if %ERRORLEVEL% neq 0 (
    echo.
    echo   [ERRO] Falha ao instalar dependencias.
    echo   Tente rodar como Administrador.
    pause
    exit /b 1
)

echo.
echo   ============================================
echo   [OK] Dependencias instaladas!
echo   ============================================
echo.
echo   Criando atalho na Area de Trabalho...
powershell -Command "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut([Environment]::GetFolderPath('Desktop') + '\FlowTrader Pro.lnk'); $s.TargetPath = '%~dp0iniciar_flowtrader.bat'; $s.WorkingDirectory = '%~dp0.'; $s.Description = 'FlowTrader Pro v3'; $s.Save()"
if %ERRORLEVEL% equ 0 (
    echo   [OK] Atalho "FlowTrader Pro" criado na Area de Trabalho!
) else (
    echo   [!] Nao foi possivel criar o atalho automaticamente.
    echo       Crie manualmente: clique direito em iniciar_flowtrader.bat
    echo       e selecione "Enviar para > Area de Trabalho"
)
echo.
echo   ============================================
echo   [OK] Setup concluido com sucesso!
echo   ============================================
echo.
echo   Para iniciar o FlowTrader:
echo     1. Abra o ProfitChart e o Excel com a planilha DDE
echo     2. Clique duas vezes no atalho "FlowTrader Pro" na Area de Trabalho
echo     3. No dashboard, abra Configuracoes para ajustar o nome do workbook
echo.
pause
