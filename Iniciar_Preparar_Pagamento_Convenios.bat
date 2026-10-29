@echo off
setlocal EnableExtensions DisableDelayedExpansion
title PREPARAR PAGAMENTO
color 0F
cd /d "%~dp0"
cls
for /F %%a in ('echo prompt $E ^| cmd') do set "ESC=%%a"
set "RESET=%ESC%[0m"
set "CYAN=%ESC%[96m"
set "BLUE=%ESC%[94m"
set "PINK=%ESC%[95m"
set "GREEN=%ESC%[92m"
set "YELLOW=%ESC%[93m"
set "WHITE=%ESC%[97m"
echo.
echo          %CYAN%SSSSS%BLUE%  EEEEE  %CYAN%SSSSS%BLUE%  DDDD   %CYAN%EEEEE%BLUE%   CCCCC%RESET%
echo          %CYAN%SS%BLUE%     EE     %CYAN%SS%BLUE%     DD  DD  %CYAN%EE%BLUE%     CC%RESET%
echo          %CYAN%SSSSS%BLUE%  EEEE   %CYAN%SSSSS%BLUE%  DD  DD  %CYAN%EEEE%BLUE%   CC%RESET%
echo          %BLUE%   SS  EE        SS  DD  DD  EE     CC%RESET%
echo          %CYAN%SSSSS%BLUE%  EEEEE  %CYAN%SSSSS%BLUE%  DDDD   %CYAN%EEEEE%BLUE%   CCCCC%RESET%
echo.
echo              %WHITE%REGULARIZACAO CONVENIOS%RESET%
echo          %PINK%Feito com ^<3 por Izabel Galvao :3%RESET%
echo          %PINK%Atualizacoes no meu GITHUB: https://github.com/i3b3l%RESET%
echo.
echo   %BLUE%^>_ %CYAN%INICIAR_PREPARAR_PAGAMENTO%RESET%
echo.
set "SCRIPT=Preparar_Pagamento_Convenios.py"
:Verificar
if not exist "%SCRIPT%" (
 echo %PINK%[ERRO]%RESET% Arquivo nao encontrado: %SCRIPT%
 goto Falha
)
set "USAR_PY="
set "PYTHON="
if exist "%~dp0.venv\Scripts\python.exe" (
 set "PYTHON=%~dp0.venv\Scripts\python.exe"
 goto PythonLocal
)
py -3 -c "import sys; sys.exit(0 if sys.version_info >= (3,10) else 1)" >nul 2>&1
if not errorlevel 1 (
 set "USAR_PY=1"
 goto Dependencias
)
python -c "import sys; sys.exit(0 if sys.version_info >= (3,10) else 1)" >nul 2>&1
if errorlevel 1 (
 echo %PINK%[ERRO]%RESET% Instale Python 3.10 ou superior e habilite a opcao PATH.
 goto Falha
)
set "PYTHON=python"
:PythonLocal
"%PYTHON%" -c "import sys; sys.exit(0 if sys.version_info >= (3,10) else 1)" >nul 2>&1
if errorlevel 1 (
 echo %PINK%[ERRO]%RESET% O Python da pasta nao funciona. Confira a instalacao.
 goto Falha
)
:Dependencias
if defined USAR_PY (
 py -3 -c "import openpyxl, keyboard, pyautogui, playwright.sync_api" >nul 2>&1
) else (
 "%PYTHON%" -c "import openpyxl, keyboard, pyautogui, playwright.sync_api" >nul 2>&1
)
if errorlevel 1 (
 echo %PINK%[ERRO]%RESET% Faltam dependencias do robo.
 echo Instale neste Python: openpyxl playwright keyboard pyautogui
 goto Falha
)
echo.
echo %BLUE%[OK]%RESET% Python e dependencias encontrados.
powershell.exe -NoProfile -Command "try { $r=Invoke-RestMethod 'http://127.0.0.1:9222/json/version' -TimeoutSec 2; if ($r.Browser -match 'Chrome/' -and $r.webSocketDebuggerUrl) { exit 0 }; exit 1 } catch { exit 1 }" >nul 2>&1
if not errorlevel 1 goto Login
set "CHROME=%ProgramFiles%\Google\Chrome\Application\chrome.exe"
if exist "%CHROME%" goto AbrirChrome
set "CHROME=%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"
if exist "%CHROME%" goto AbrirChrome
set "CHROME=%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"
if not exist "%CHROME%" (
 echo %PINK%[ERRO]%RESET% Google Chrome nao encontrado.
 goto Falha
)
:AbrirChrome
echo %CYAN%[CHROME]%RESET% Abrindo o navegador para login no SIGEF...
start "Chrome SIGEF" "%CHROME%" --remote-debugging-port=9222 --remote-debugging-address=127.0.0.1 --user-data-dir="%LOCALAPPDATA%\IzabelSIGEF\ChromeDebugProfile" "https://sigef.sefin.ro.gov.br"
:Login
echo.
echo %CYAN%+----------------------------------------------------------+%RESET%
echo %CYAN%^|%RESET% %WHITE%PREPARAR PAGAMENTOS%RESET%                                      %CYAN%^|%RESET%
echo %CYAN%+----------------------------------------------------------+%RESET%
echo   O robo ira consultar as NLs pendentes, conferir o credor e
echo   o domicilio bancario e salvar as PPs geradas na planilha.
echo.
echo   %GREEN%Quando estiver pronto, pressione qualquer tecla.%RESET%
pause
echo.
echo %BLUE%[INICIANDO]%RESET% %SCRIPT%
echo.
if defined USAR_PY (
 py -3 -u "%SCRIPT%"
) else (
 "%PYTHON%" -u "%SCRIPT%"
)
if errorlevel 1 goto Falha
echo.
echo %GREEN%[ENCERRADO]%RESET% Confira o resumo da execucao acima.
echo.
pause
exit /b 0
:Falha
echo.
echo %YELLOW%[ATENCAO]%RESET% A execucao nao foi concluida normalmente.
echo Confira as mensagens acima antes de tentar novamente.
pause
exit /b 1
