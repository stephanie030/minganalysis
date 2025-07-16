@echo off
chcp 65001
echo ========================================
echo 設定每日庫存分析自動執行任務
echo ========================================

echo.
echo 正在設定Windows工作排程器任務...
echo.

REM 刪除現有任務（如果存在）
schtasks /delete /tn "每日庫存分析" /f >nul 2>&1

REM 創建新的每日任務，每天早上8點執行
schtasks /create /tn "每日庫存分析" /tr "python \"%~dp0daily_analysis.py\"" /sc daily /st 08:00 /ru SYSTEM

if %errorlevel% equ 0 (
    echo ✅ 任務設定成功！
    echo.
    echo 📅 任務名稱: 每日庫存分析
    echo ⏰ 執行時間: 每天早上 8:00
    echo 📁 執行目錄: %~dp0
    echo.
    echo 您可以在「工作排程器」中查看和管理此任務
    echo.
) else (
    echo ❌ 任務設定失敗！
    echo 請以管理員身份執行此批次檔案
    echo.
)

echo 按任意鍵繼續...
pause >nul
