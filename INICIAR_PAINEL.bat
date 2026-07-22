@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================
echo    PAINEL DE PRODUTIVIDADE - ATIVOS E MP
echo ============================================
echo.
echo [1/2] Instalando dependencias (so na primeira vez)...
python -m pip install --quiet --upgrade streamlit pandas plotly openpyxl
if errorlevel 1 (
  echo.
  echo *** Nao consegui rodar o Python. Instale em https://www.python.org/downloads/
  echo *** e marque "Add Python to PATH". Depois rode este arquivo de novo.
  pause
  exit /b
)
echo [2/2] Abrindo o painel no navegador...
echo.
echo (Para FECHAR o painel depois: volte nesta janela e aperte Ctrl+C)
python -m streamlit run painel_produtividade.py
pause
