@echo off
echo Starting Houzz Converter App...
cd /d "%~dp0"
streamlit run app.py
pause
