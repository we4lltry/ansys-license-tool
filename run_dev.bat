@echo off
chcp 65001 > nul
echo [테스트] 앱을 직접 실행합니다...
IF EXIST ".venv\Scripts\activate.bat" call .venv\Scripts\activate.bat
python -m streamlit run app.py --server.port 8501 --theme.base dark
pause
