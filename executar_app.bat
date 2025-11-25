@echo off
echo ========================================
echo   Sistema de Convocacao - TCE-MA
echo   Iniciando painel interativo...
echo ========================================
echo.

REM Ativar ambiente virtual
call .venv\Scripts\activate

REM Executar app Streamlit
streamlit run app.py

pause
