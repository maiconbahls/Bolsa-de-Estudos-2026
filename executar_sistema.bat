@echo off
chcp 65001 > nul
echo ========================================================
echo   INICIANDO SISTEMA DE BOLSAS DE ESTUDOS
echo ========================================================
echo.

:: Verifica se o Python está instalado
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERRO] Python nao encontrado!
    echo Por favor, instale o Python em https://www.python.org/downloads/
    echo Lembre-se de marcar a opcao "Add Python to PATH" durante a instalacao.
    pause
    exit /b
)

echo [1/2] Verificando bibliotecas necessarias...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo [ERRO] Falha ao instalar as bibliotecas. Verifique sua conexao com a internet.
    pause
    exit /b
)

echo.
echo [2/2] Iniciando aplicativo...
echo Pressione CTRL+C para encerrar.
echo.

streamlit run app.py

pause
