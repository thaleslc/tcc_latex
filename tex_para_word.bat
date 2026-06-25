@echo off
setlocal

cd /d "%~dp0"

set "MAIN=main.tex"
set "OUTDIR=output\word"
set "OUTDOCX=%OUTDIR%\tcc_convertido_melhorado.docx"
set "FALLBACK=scripts\latex_to_docx_simple.py"
set "REFERENCE_DOC=templates\pandoc_reference_tcc.docx"
set "REFERENCE_BUILDER=scripts\create_tcc_reference_docx.py"
set "POSTPROCESS=scripts\postprocess_tcc_docx.py"
set "BUNDLED_PYTHON=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"

if not exist "%MAIN%" (
    echo Erro: %MAIN% nao foi encontrado nesta pasta.
    pause
    exit /b 1
)

if not exist "%OUTDIR%" mkdir "%OUTDIR%"

set "PYTHON="
if exist "%BUNDLED_PYTHON%" (
    set "PYTHON=%BUNDLED_PYTHON%"
) else (
    where python >nul 2>nul
    if not errorlevel 1 set "PYTHON=python"
)

if not "%PYTHON%"=="" if exist "%REFERENCE_BUILDER%" (
    "%PYTHON%" "%REFERENCE_BUILDER%"
)

set "PANDOC="
for /r "%CD%" %%F in (pandoc.exe) do if not defined PANDOC if exist "%%~fF" set "PANDOC=%%~fF"
if "%PANDOC%"=="" (
    where pandoc >nul 2>nul
    if not errorlevel 1 set "PANDOC=pandoc"
)
if "%PANDOC%"=="" if exist "%ProgramFiles%\Pandoc\pandoc.exe" set "PANDOC=%ProgramFiles%\Pandoc\pandoc.exe"
if "%PANDOC%"=="" if exist "%LOCALAPPDATA%\Pandoc\pandoc.exe" set "PANDOC=%LOCALAPPDATA%\Pandoc\pandoc.exe"

if not "%PANDOC%"=="" (
    echo.
    echo Convertendo %MAIN% para Word com Pandoc...
    echo Usando: %PANDOC%
    echo.
    if exist "%REFERENCE_DOC%" (
        "%PANDOC%" "%MAIN%" --from=latex --to=docx --standalone --resource-path=".;figuras" --reference-doc="%REFERENCE_DOC%" --output="%OUTDOCX%"
    ) else (
        "%PANDOC%" "%MAIN%" --from=latex --to=docx --standalone --resource-path=".;figuras" --output="%OUTDOCX%"
    )
    if not errorlevel 1 goto success

    echo.
    echo Pandoc falhou. Tentando conversor simples em Python...
)

if "%PYTHON%"=="" (
        echo.
        echo Erro: nem Pandoc nem Python foram encontrados.
        echo Instale o Pandoc para a melhor conversao LaTeX para Word:
        echo https://pandoc.org/installing.html
        pause
        exit /b 1
)

if not exist "%FALLBACK%" (
    echo Erro: conversor fallback nao encontrado em %FALLBACK%.
    pause
    exit /b 1
)

echo.
echo Convertendo %MAIN% para Word com conversor simples em Python...
echo.
"%PYTHON%" "%FALLBACK%" "%MAIN%" "%OUTDOCX%"
if errorlevel 1 (
    echo.
    echo Erro durante a conversao para Word.
    echo Feche o arquivo Word se ele estiver aberto e tente novamente.
    pause
    exit /b 1
)

:success
if not "%PYTHON%"=="" if exist "%POSTPROCESS%" (
    echo.
    echo Aplicando formatacao do TCC no Word...
    "%PYTHON%" "%POSTPROCESS%" "%OUTDOCX%"
    if errorlevel 1 (
        echo.
        echo Aviso: o Word foi gerado, mas o pos-processamento de formato falhou.
        echo O arquivo ainda esta disponivel em %OUTDOCX%.
        pause
        exit /b 1
    )
)

echo.
echo Word gerado com sucesso:
echo %CD%\%OUTDOCX%
echo.
echo Abra esse arquivo no Word para revisar.
exit /b 0
