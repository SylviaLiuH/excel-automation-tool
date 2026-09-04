@echo off
chcp 65001 >nul

echo ================================
echo   Excel 自动整理器启动中...
echo ================================

py -m pip install -r requirements.txt

echo.
echo 正在启动网页...
py -m streamlit run app.py

pause