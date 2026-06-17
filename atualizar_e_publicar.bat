@echo off
cd /d "%~dp0"
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
    pause
    exit /b 1
)

echo.
echo [2/2] Publicando no GitHub...
git add -A
git commit -m "atualiza dados"
echo.
echo [2a/2] Puxando commits remotos...
git pull origin main
if errorlevel 1 (
    echo.
    echo [ERRO] Falha ao puxar commits remotos.
    echo Solucao: abra Git Bash e rode "git pull origin main" manualmente.
    pause
    exit /b 1
)
echo.
echo [2b/2] Enviando para GitHub...
git push origin main

echo.
echo ========================================
echo   PRONTO! Dashboard atualizado em ~1min:
echo   https://grdmmplan-source.github.io/dashboard-firjan/
echo ========================================
echo.
pause
