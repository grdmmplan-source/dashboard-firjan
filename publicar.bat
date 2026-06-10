@echo off
cd /d "%~dp0"
echo.
echo ========================================
echo   PUBLICANDO DASHBOARD FIRJAN...
echo ========================================
echo.

git add index.html
git commit -m "atualiza dados"
echo.
echo [1/2] Puxando commits remotos...
git pull origin main
if errorlevel 1 (
    echo.
    echo [ERRO] Falha ao puxar commits remotos.
    echo Solucao: abra Git Bash e rode "git pull origin main" manualmente.
    pause
    exit /b 1
)
echo.
echo [2/2] Enviando para GitHub...
git push origin main

echo.
echo ========================================
echo   PRONTO! Link atualizado em ~1 minuto:
echo   https://grdmmplan-source.github.io/dashboard-firjan/
echo ========================================
echo.
pause
