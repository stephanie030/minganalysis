@echo off
chcp 65001
echo ========================================
echo 啟動庫存分析動態儀表板
echo ========================================

echo.
echo 正在啟動 Streamlit 儀表板...
echo 請稍候，瀏覽器將自動開啟
echo.

REM 激活虛擬環境並安裝 streamlit 和 plotly
call .venv\Scripts\activate
uv pip install streamlit plotly

REM 啟動 Streamlit 應用程式
streamlit run streamlit_dashboard.py

echo.
echo 儀表板已關閉
pause
