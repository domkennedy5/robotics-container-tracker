@echo off
cd /d "%~dp0"
start "" "http://localhost:8502"
python -m streamlit run app.py --server.port 8502 --server.headless false
