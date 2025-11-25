@echo off
echo ========================================
echo Sistema de Geracao de Lista de Convocacao
echo TCE-MA - SUDEC
echo ========================================
echo.

REM Ativar ambiente virtual e executar o script
call .venv\Scripts\activate.bat
python gerar_lista_convocacao.py

echo.
echo ========================================
echo Pressione qualquer tecla para sair...
pause > nul
