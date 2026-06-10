@echo off
cd /d "%~dp0"
echo.
echo ========================================
echo   FIRJAN — Limpar todos os dados
echo ========================================
echo.

set /p confirm="Tem certeza? Isso apaga todos os dados do dashboard. (S/N): "
if /i "%confirm%" neq "S" (
    echo Operacao cancelada.
    pause
    exit /b 0
)

echo.
echo Restaurando dashboard zerado...
copy /y "_template_zerado.html" "index.html" > nul

echo.
echo ========================================
echo   PRONTO! Todos os dados foram limpos.
echo   Rode atualizar.bat para repopular.
echo ========================================
echo.
pause
