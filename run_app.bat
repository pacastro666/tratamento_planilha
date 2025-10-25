@echo off
echo Ativando ambiente virtual...
call venv\Scripts\activate
echo.
echo Iniciando aplicativo Streamlit...
echo O aplicativo será aberto automaticamente no navegador.
echo Para parar o aplicativo, pressione Ctrl+C
echo.
streamlit run app.py
