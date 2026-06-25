@echo off
setlocal

cd /d "%~dp0"

set "MAIN=main.tex"
set "PDF=main.pdf"
set "OUTDIR=output\pdf"
set "OUTPDF=%OUTDIR%\tcc_renderizado.pdf"
set "MIKTEX_XELATEX=%LOCALAPPDATA%\Programs\MiKTeX\miktex\bin\x64\xelatex.exe"

if exist "%MIKTEX_XELATEX%" (
    set "XELATEX=%MIKTEX_XELATEX%"
) else (
    where xelatex >nul 2>nul
    if errorlevel 1 (
        echo Erro: xelatex nao foi encontrado.
        echo Instale o MiKTeX ou adicione o xelatex ao PATH.
        pause
        exit /b 1
    )
    set "XELATEX=xelatex"
)

if not exist "%MAIN%" (
    echo Erro: %MAIN% nao foi encontrado nesta pasta.
    pause
    exit /b 1
)

echo.
echo Compilando %MAIN% com XeLaTeX...
echo.

"%XELATEX%" -file-line-error -interaction=nonstopmode -halt-on-error "%MAIN%"
if errorlevel 1 goto compile_error

"%XELATEX%" -file-line-error -interaction=nonstopmode -halt-on-error "%MAIN%"
if errorlevel 1 goto compile_error

if not exist "%PDF%" (
    echo Erro: a compilacao terminou, mas %PDF% nao foi gerado.
    pause
    exit /b 1
)

if not exist "%OUTDIR%" mkdir "%OUTDIR%"
copy /Y "%PDF%" "%OUTPDF%" >nul
if errorlevel 1 (
    echo Erro: nao foi possivel copiar o PDF para %OUTPDF%.
    echo Feche o PDF se ele estiver aberto em um visualizador que bloqueia o arquivo.
    pause
    exit /b 1
)

echo.
echo PDF gerado com sucesso:
echo %CD%\%OUTPDF%
echo.

start "" "%OUTPDF%"
exit /b 0

:compile_error
echo.
echo Erro durante a compilacao.
echo Veja as mensagens acima ou abra main.log para detalhes.
pause
exit /b 1
