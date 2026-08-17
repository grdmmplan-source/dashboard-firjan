@echo off
setlocal enabledelayedexpansion
REM ==========================================================
REM  Versao para Tarefa Agendada (Task Scheduler) - sem pause,
REM  sem interacao.
REM
REM  IMPORTANTE: todo comando git roda com "< NUL" para que o
REM  stdin nunca seja um console interativo. Isso evita que o
REM  Git for Windows trave esperando resposta a prompts como
REM  "Rename from X to Y failed. Should I try again? (y/n)"
REM  quando o .git esta num compartilhamento de rede (SMB) e
REM  o rename de um pack/idx falha momentaneamente. Sem essa
REM  trava, o processo nao fica pendurado e as proximas
REM  execucoes agendadas (11h, 12h, ...) nao ficam bloqueadas
REM  esperando essa instancia anterior terminar.
REM ==========================================================
pushd "%~dp0"
echo.
echo ========================================
echo   FIRJAN — Atualizar e Publicar
echo ========================================
echo.

echo [1/2] Calculando dados do Excel...
python atualizar_tudo.py --no-pause
if %errorlevel% neq 0 (
    echo.
    echo [ERRO] Falha ao processar o Excel.
    exit /b 1
)

echo.
echo [2/2] Publicando no GitHub...
git add -A < NUL
git commit -m "atualiza dados" < NUL

echo.
echo [2a/2] Puxando commits remotos (com retry automatico)...
set PULL_OK=0
for /L %%i in (1,1,3) do (
    if "!PULL_OK!"=="0" (
        git pull origin main < NUL
        if not errorlevel 1 (
            set PULL_OK=1
        ) else (
            echo [AVISO] Tentativa %%i de pull falhou. Tentando novamente em 5s...
            timeout /t 5 /nobreak < NUL >nul
        )
    )
)
if "!PULL_OK!"=="0" (
    echo.
    echo [ERRO] Falha ao puxar commits remotos apos 3 tentativas.
    echo Solucao: abra Git Bash e rode "git pull origin main" manualmente.
    exit /b 1
)

echo.
echo [2b/2] Enviando para GitHub (com retry automatico)...
set PUSH_OK=0
for /L %%i in (1,1,3) do (
    if "!PUSH_OK!"=="0" (
        git push origin main < NUL
        if not errorlevel 1 (
            set PUSH_OK=1
        ) else (
            echo [AVISO] Tentativa %%i de push falhou. Tentando novamente em 5s...
            timeout /t 5 /nobreak < NUL >nul
        )
    )
)
if "!PUSH_OK!"=="0" (
    echo.
    echo [ERRO] git push falhou apos 3 tentativas.
    echo Solucao: abra Git Bash e rode "git push origin main" manualmente.
    exit /b 1
)

echo.
echo ========================================
echo   PRONTO! Dashboard atualizado em ~1min:
echo   https://grdmmplan-source.github.io/dashboard-firjan/
echo ========================================
echo.
exit /b 0
