@echo off
setlocal
cd /d "%~dp0"

title Robo Migracao Peterson - Definitivo

echo ============================================================
echo ROBO MIGRACAO PETERSON - MODO TURBO
echo ============================================================
echo.

if not exist "credentials.json" (
    echo ERRO: credentials.json nao encontrado.
    pause
    exit /b 1
)

if not exist "token.json" (
    echo ERRO: token.json nao encontrado.
    echo Rode primeiro autenticar_google_peterson.bat
    pause
    exit /b 1
)

echo Instalando/atualizando dependencias...
py -m pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib pywinauto pywin32 pyautogui pyperclip

if errorlevel 1 (
    echo.
    echo ERRO nas dependencias.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo OPERACAO
echo ============================================================
echo 1 - Ver status
echo 2 - Corrigir coluna O
echo 3 - Teste: enviar 1
echo 4 - Enviar 5
echo 5 - Enviar toda a fila
echo 6 - Etiquetar contatos ja em andamento
echo 7 - Sair
echo.

set /p OPCAO=Escolha: 

if "%OPCAO%"=="1" py robo_migracao_peterson.py --status
if "%OPCAO%"=="2" py robo_migracao_peterson.py --corrigir-o
if "%OPCAO%"=="3" py robo_migracao_peterson.py --enviar 1
if "%OPCAO%"=="4" py robo_migracao_peterson.py --enviar 5
if "%OPCAO%"=="5" py robo_migracao_peterson.py --enviar todos
if "%OPCAO%"=="6" py robo_migracao_peterson.py --etiquetar todos
if "%OPCAO%"=="7" goto FIM

:FIM
echo.
pause
