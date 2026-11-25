@echo off
title ViabilizaLab
cd /d "%~dp0"
set "ARQ="
for /f "delims=" %%F in ('dir /b /o:-d "app_pnbox*.py" 2^>nul') do (
    set "ARQ=%%F"
    goto :achou
)
echo.
echo Nao encontrei o arquivo do programa (app_pnbox.py) nesta pasta:
echo   %CD%
echo Coloque este .bat na MESMA pasta do arquivo .py e tente de novo.
echo.
pause
exit /b 1

:achou
echo Iniciando o ViabilizaLab (%ARQ%)...
echo Ao fechar a aba do navegador, esta janela fecha sozinha.
where py >nul 2>nul
if %errorlevel%==0 (
    py "%ARQ%"
) else (
    python "%ARQ%"
)
if errorlevel 1 pause
