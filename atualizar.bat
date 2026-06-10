@echo off
cd /d "%~dp0"
echo.
echo ========================================
echo   FIRJAN — Atualizando Dashboard...
echo ========================================
echo.

python atualizar_tudo.py

echo.
echo ========================================
echo   Abra o index.html para visualizar.
echo   Quando estiver pronto, rode publicar.bat
echo   para enviar ao GitHub.
echo ========================================
echo.
pause
