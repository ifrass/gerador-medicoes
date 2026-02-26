@echo off
chcp 65001 > nul
echo ==================================================
echo 🔄 Sincronizando com a versao mais recente...
echo ==================================================

:: Puxa as atualizacoes do GitHub silenciosamente
git pull

echo.
echo 📦 Verificando dependencias (Bibliotecas base)...
python -m pip install streamlit openpyxl --quiet

echo.
echo Iniciando o extrator no navegador...
python -m streamlit run app.py

pause