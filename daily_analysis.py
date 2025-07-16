#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
每日庫存分析自動執行腳本
建議設定為每日早上8點自動執行
"""

import os
import sys
import glob
from datetime import datetime
from pathlib import Path
from inventory_analyzer import InventoryAnalyzer

def find_latest_data_file():
    """尋找最新的資料檔案"""
    # 尋找所有Excel檔案
    excel_files = glob.glob("*.xlsx") + glob.glob("*.xls")
    
    if not excel_files:
        print("❌ 找不到Excel資料檔案")
        return None
    
    # 按修改時間排序，取最新的
    latest_file = max(excel_files, key=os.path.getmtime)
    print(f"📁 找到最新資料檔案: {latest_file}")
    
    return latest_file

def archive_old_reports():
    """歸檔舊報告"""
    # 創建歷史報告目錄
    archive_dir = Path("reports_archive")
    archive_dir.mkdir(exist_ok=True)
    
    # 移動7天前的報告到歷史目錄
    import shutil
    from datetime import timedelta
    
    cutoff_date = datetime.now() - timedelta(days=7)
    
    for report_file in glob.glob("庫存分析報告_*.md"):
        file_time = datetime.fromtimestamp(os.path.getmtime(report_file))
        if file_time < cutoff_date:
            shutil.move(report_file, archive_dir / report_file)
            print(f"📦 歸檔舊報告: {report_file}")

def send_notification(report_file, success=True):
    """發送通知（可擴展為郵件、Slack等）"""
    if success:
        print(f"✅ 分析完成通知: 報告已生成 - {report_file}")
        # 這裡可以添加郵件發送、Slack通知等功能
    else:
        print("❌ 分析失敗通知: 請檢查資料檔案和程式")

def main():
    """主程式"""
    print("="*60)
    print(f"🕐 每日庫存分析 - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("="*60)
    
    try:
        # 1. 尋找最新資料檔案
        data_file = find_latest_data_file()
        if not data_file:
            send_notification("", False)
            return False
        
        # 2. 歸檔舊報告
        archive_old_reports()
        
        # 3. 執行分析
        analyzer = InventoryAnalyzer(data_file)
        success = analyzer.run_full_analysis()
        
        if success:
            # 4. 發送成功通知
            report_files = glob.glob("庫存分析報告_*.md")
            latest_report = max(report_files, key=os.path.getmtime) if report_files else "未知"
            send_notification(latest_report, True)
            
            print("\n" + "="*60)
            print("🎯 每日分析任務完成！")
            print("="*60)
            return True
        else:
            send_notification("", False)
            return False
            
    except Exception as e:
        print(f"❌ 執行過程中發生錯誤: {e}")
        send_notification("", False)
        return False

if __name__ == "__main__":
    main()
