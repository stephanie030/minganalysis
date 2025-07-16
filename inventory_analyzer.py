#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
庫存資料分析系統
每日自動分析庫存變化趨勢、異常監控和報告生成
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime, timedelta
import warnings
import os
from pathlib import Path

# 設定中文字體
plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei', 'SimHei']
plt.rcParams['axes.unicode_minus'] = False
warnings.filterwarnings('ignore')

class InventoryAnalyzer:
    def __init__(self, data_file=None):
        """初始化分析器"""
        self.data_file = data_file or "銘宣爬蟲合併_0421_0708.xlsx"
        self.df = None
        self.analysis_results = {}
        
    def load_data(self):
        """載入資料"""
        try:
            print(f"📊 載入資料檔案: {self.data_file}")
            self.df = pd.read_excel(self.data_file)
            
            # 資料清理
            self.df['爬取日期'] = pd.to_datetime(self.df['爬取日期'])
            self.df['客戶庫存量(現有數量)'] = pd.to_numeric(self.df['客戶庫存量(現有數量)'], errors='coerce')
            self.df['包裝_數量'] = pd.to_numeric(self.df['包裝_數量'], errors='coerce')
            
            print(f"✅ 資料載入成功！總筆數: {len(self.df):,}")
            return True
        except Exception as e:
            print(f"❌ 資料載入失敗: {e}")
            return False
    
    def basic_statistics(self):
        """基本統計分析"""
        print("\n" + "="*50)
        print("📈 基本統計分析")
        print("="*50)
        
        stats = {
            '總筆數': len(self.df),
            '料號數量': self.df['客戶料號'].nunique(),
            '產品種類': self.df['客戶品名'].nunique(),
            '倉庫數量': self.df['客戶庫別'].nunique(),
            '資料天數': self.df['爬取日期'].nunique(),
            '時間範圍': f"{self.df['爬取日期'].min().strftime('%Y-%m-%d')} 到 {self.df['爬取日期'].max().strftime('%Y-%m-%d')}"
        }
        
        for key, value in stats.items():
            print(f"{key}: {value:,}" if isinstance(value, int) else f"{key}: {value}")
        
        self.analysis_results['basic_stats'] = stats
        return stats
    
    def warehouse_analysis(self):
        """倉庫分析"""
        print("\n" + "="*50)
        print("🏭 倉庫庫存分析")
        print("="*50)
        
        # 各倉庫庫存統計
        warehouse_stats = self.df.groupby('客戶庫別').agg({
            '客戶庫存量(現有數量)': ['sum', 'mean', 'count'],
            '客戶料號': 'nunique'
        }).round(2)
        
        warehouse_stats.columns = ['總庫存量', '平均庫存量', '品項數量', '料號種類']
        warehouse_stats = warehouse_stats.sort_values('總庫存量', ascending=False)
        
        print("各倉庫庫存概況:")
        print(warehouse_stats)
        
        self.analysis_results['warehouse_stats'] = warehouse_stats
        return warehouse_stats
    
    def trend_analysis(self):
        """趨勢分析"""
        print("\n" + "="*50)
        print("📊 庫存趨勢分析")
        print("="*50)
        
        # 每日總庫存趨勢
        daily_inventory = self.df.groupby('爬取日期')['客戶庫存量(現有數量)'].sum().reset_index()
        daily_inventory['日期'] = daily_inventory['爬取日期'].dt.strftime('%Y-%m-%d')
        
        print("每日總庫存量:")
        for _, row in daily_inventory.iterrows():
            print(f"{row['日期']}: {row['客戶庫存量(現有數量)']:,.0f}")
        
        # 計算變化率
        if len(daily_inventory) > 1:
            daily_inventory['變化量'] = daily_inventory['客戶庫存量(現有數量)'].diff()
            daily_inventory['變化率(%)'] = (daily_inventory['變化量'] / daily_inventory['客戶庫存量(現有數量)'].shift(1) * 100).round(2)
            
            print("\n每日庫存變化:")
            for _, row in daily_inventory[1:].iterrows():
                change_symbol = "📈" if row['變化量'] > 0 else "📉" if row['變化量'] < 0 else "➡️"
                print(f"{row['日期']}: {change_symbol} {row['變化量']:+,.0f} ({row['變化率(%)']:+.1f}%)")
        
        self.analysis_results['daily_trend'] = daily_inventory
        return daily_inventory
    
    def product_analysis(self):
        """產品分析"""
        print("\n" + "="*50)
        print("🔍 產品庫存分析")
        print("="*50)
        
        # 最新日期的產品庫存
        latest_date = self.df['爬取日期'].max()
        latest_data = self.df[self.df['爬取日期'] == latest_date]
        
        # 各產品總庫存
        product_inventory = latest_data.groupby(['客戶料號', '客戶品名']).agg({
            '客戶庫存量(現有數量)': 'sum',
            '客戶庫別': 'count'
        }).reset_index()
        
        product_inventory.columns = ['客戶料號', '客戶品名', '總庫存量', '倉庫數量']
        product_inventory = product_inventory.sort_values('總庫存量', ascending=False)
        
        print("庫存量前10名產品:")
        print(product_inventory.head(10)[['客戶料號', '客戶品名', '總庫存量']].to_string(index=False))
        
        # 零庫存產品
        zero_inventory = product_inventory[product_inventory['總庫存量'] == 0]
        if len(zero_inventory) > 0:
            print(f"\n⚠️  零庫存產品數量: {len(zero_inventory)}")
            if len(zero_inventory) <= 10:
                print("零庫存產品:")
                print(zero_inventory[['客戶料號', '客戶品名']].to_string(index=False))
        
        self.analysis_results['product_stats'] = product_inventory
        self.analysis_results['zero_inventory'] = zero_inventory
        return product_inventory
    
    def anomaly_detection(self):
        """異常檢測"""
        print("\n" + "="*50)
        print("🚨 異常檢測")
        print("="*50)
        
        anomalies = []
        
        # 檢測庫存異常變化
        if len(self.df['爬取日期'].unique()) > 1:
            for product in self.df['客戶料號'].unique():
                product_data = self.df[self.df['客戶料號'] == product].groupby('爬取日期')['客戶庫存量(現有數量)'].sum().reset_index()
                
                if len(product_data) > 1:
                    product_data['變化率'] = product_data['客戶庫存量(現有數量)'].pct_change() * 100
                    
                    # 檢測大幅變化 (>50%)
                    large_changes = product_data[abs(product_data['變化率']) > 50]
                    
                    for _, row in large_changes.iterrows():
                        if not pd.isna(row['變化率']):
                            anomalies.append({
                                '料號': product,
                                '日期': row['爬取日期'].strftime('%Y-%m-%d'),
                                '類型': '庫存大幅變化',
                                '變化率': f"{row['變化率']:+.1f}%",
                                '當前庫存': row['客戶庫存量(現有數量)']
                            })
        
        if anomalies:
            print(f"發現 {len(anomalies)} 個異常:")
            for anomaly in anomalies[:10]:  # 只顯示前10個
                print(f"• {anomaly['料號']} ({anomaly['日期']}): {anomaly['類型']} {anomaly['變化率']}")
        else:
            print("✅ 未發現明顯異常")
        
        self.analysis_results['anomalies'] = anomalies
        return anomalies
    
    def generate_visualizations(self):
        """生成視覺化圖表"""
        print("\n" + "="*50)
        print("📊 生成視覺化圖表")
        print("="*50)
        
        # 創建圖表目錄
        charts_dir = Path("charts")
        charts_dir.mkdir(exist_ok=True)
        
        # 1. 每日庫存趨勢圖
        if 'daily_trend' in self.analysis_results:
            plt.figure(figsize=(12, 6))
            daily_data = self.analysis_results['daily_trend']
            plt.plot(daily_data['爬取日期'], daily_data['客戶庫存量(現有數量)'], marker='o', linewidth=2, markersize=8)
            plt.title('每日總庫存量趨勢', fontsize=16, fontweight='bold')
            plt.xlabel('日期', fontsize=12)
            plt.ylabel('庫存量', fontsize=12)
            plt.xticks(rotation=45)
            plt.grid(True, alpha=0.3)
            plt.tight_layout()
            plt.savefig(charts_dir / '每日庫存趨勢.png', dpi=300, bbox_inches='tight')
            plt.close()
            print("✅ 已生成: 每日庫存趨勢.png")
        
        # 2. 倉庫庫存分布圖
        if 'warehouse_stats' in self.analysis_results:
            plt.figure(figsize=(10, 6))
            warehouse_data = self.analysis_results['warehouse_stats']
            plt.bar(warehouse_data.index, warehouse_data['總庫存量'])
            plt.title('各倉庫庫存分布', fontsize=16, fontweight='bold')
            plt.xlabel('倉庫', fontsize=12)
            plt.ylabel('總庫存量', fontsize=12)
            plt.xticks(rotation=45)
            plt.grid(True, alpha=0.3, axis='y')
            plt.tight_layout()
            plt.savefig(charts_dir / '倉庫庫存分布.png', dpi=300, bbox_inches='tight')
            plt.close()
            print("✅ 已生成: 倉庫庫存分布.png")
        
        return True

    def generate_report(self):
        """生成分析報告"""
        print("\n" + "="*50)
        print("📋 生成分析報告")
        print("="*50)

        report_content = []
        report_content.append("# 庫存分析報告")
        report_content.append(f"**生成時間**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        report_content.append("")

        # 基本統計
        if 'basic_stats' in self.analysis_results:
            report_content.append("## 📊 基本統計")
            stats = self.analysis_results['basic_stats']
            for key, value in stats.items():
                report_content.append(f"- **{key}**: {value:,}" if isinstance(value, int) else f"- **{key}**: {value}")
            report_content.append("")

        # 倉庫分析
        if 'warehouse_stats' in self.analysis_results:
            report_content.append("## 🏭 倉庫分析")
            warehouse_data = self.analysis_results['warehouse_stats']
            report_content.append("| 倉庫 | 總庫存量 | 平均庫存量 | 品項數量 | 料號種類 |")
            report_content.append("|------|----------|------------|----------|----------|")
            for warehouse, row in warehouse_data.iterrows():
                report_content.append(f"| {warehouse} | {row['總庫存量']:,.0f} | {row['平均庫存量']:,.1f} | {row['品項數量']:,} | {row['料號種類']:,} |")
            report_content.append("")

        # 趨勢分析
        if 'daily_trend' in self.analysis_results:
            report_content.append("## 📈 庫存趨勢")
            daily_data = self.analysis_results['daily_trend']
            report_content.append("| 日期 | 總庫存量 | 變化量 | 變化率(%) |")
            report_content.append("|------|----------|--------|-----------|")
            for _, row in daily_data.iterrows():
                change_str = f"{row['變化量']:+,.0f}" if not pd.isna(row['變化量']) else "-"
                rate_str = f"{row['變化率(%)']:+.1f}%" if not pd.isna(row['變化率(%)']) else "-"
                report_content.append(f"| {row['日期']} | {row['客戶庫存量(現有數量)']:,.0f} | {change_str} | {rate_str} |")
            report_content.append("")

        # 產品分析
        if 'product_stats' in self.analysis_results:
            report_content.append("## 🔍 產品分析")
            product_data = self.analysis_results['product_stats']

            report_content.append("### 庫存量前10名產品")
            report_content.append("| 料號 | 產品名稱 | 總庫存量 | 倉庫數量 |")
            report_content.append("|------|----------|----------|----------|")
            for _, row in product_data.head(10).iterrows():
                report_content.append(f"| {row['客戶料號']} | {row['客戶品名']} | {row['總庫存量']:,.0f} | {row['倉庫數量']} |")
            report_content.append("")

            # 零庫存警告
            if 'zero_inventory' in self.analysis_results and len(self.analysis_results['zero_inventory']) > 0:
                zero_count = len(self.analysis_results['zero_inventory'])
                report_content.append(f"### ⚠️ 零庫存警告 ({zero_count}個產品)")
                if zero_count <= 20:
                    report_content.append("| 料號 | 產品名稱 |")
                    report_content.append("|------|----------|")
                    for _, row in self.analysis_results['zero_inventory'].iterrows():
                        report_content.append(f"| {row['客戶料號']} | {row['客戶品名']} |")
                else:
                    report_content.append(f"零庫存產品過多({zero_count}個)，請檢查詳細清單。")
                report_content.append("")

        # 異常檢測
        if 'anomalies' in self.analysis_results:
            anomalies = self.analysis_results['anomalies']
            if anomalies:
                report_content.append(f"## 🚨 異常檢測 (發現{len(anomalies)}個異常)")
                report_content.append("| 料號 | 日期 | 異常類型 | 變化率 | 當前庫存 |")
                report_content.append("|------|------|----------|--------|----------|")
                for anomaly in anomalies[:20]:  # 最多顯示20個
                    report_content.append(f"| {anomaly['料號']} | {anomaly['日期']} | {anomaly['類型']} | {anomaly['變化率']} | {anomaly['當前庫存']:,.0f} |")
                if len(anomalies) > 20:
                    report_content.append(f"... 還有 {len(anomalies)-20} 個異常未顯示")
            else:
                report_content.append("## ✅ 異常檢測")
                report_content.append("未發現明顯異常。")
            report_content.append("")

        # 建議
        report_content.append("## 💡 建議")
        suggestions = self._generate_suggestions()
        for suggestion in suggestions:
            report_content.append(f"- {suggestion}")

        # 保存報告
        report_file = f"庫存分析報告_{datetime.now().strftime('%Y%m%d_%H%M%S')}.md"
        with open(report_file, 'w', encoding='utf-8') as f:
            f.write('\n'.join(report_content))

        print(f"✅ 報告已保存: {report_file}")
        return report_file

    def _generate_suggestions(self):
        """生成建議"""
        suggestions = []

        # 基於分析結果生成建議
        if 'zero_inventory' in self.analysis_results:
            zero_count = len(self.analysis_results['zero_inventory'])
            if zero_count > 0:
                suggestions.append(f"有 {zero_count} 個產品庫存為零，建議及時補貨")

        if 'anomalies' in self.analysis_results:
            anomaly_count = len(self.analysis_results['anomalies'])
            if anomaly_count > 0:
                suggestions.append(f"發現 {anomaly_count} 個庫存異常變化，建議進一步調查原因")

        if 'warehouse_stats' in self.analysis_results:
            warehouse_data = self.analysis_results['warehouse_stats']
            if len(warehouse_data) > 1:
                max_warehouse = warehouse_data['總庫存量'].idxmax()
                min_warehouse = warehouse_data['總庫存量'].idxmin()
                suggestions.append(f"庫存分布不均，{max_warehouse}庫存最多，{min_warehouse}庫存最少，考慮調撥平衡")

        if 'daily_trend' in self.analysis_results:
            daily_data = self.analysis_results['daily_trend']
            if len(daily_data) > 1:
                latest_change = daily_data.iloc[-1]['變化率(%)']
                if not pd.isna(latest_change):
                    if latest_change < -10:
                        suggestions.append("最近庫存下降超過10%，建議關注補貨計劃")
                    elif latest_change > 20:
                        suggestions.append("最近庫存增長超過20%，建議檢查是否有積壓風險")

        if not suggestions:
            suggestions.append("目前庫存狀況良好，建議持續監控")

        return suggestions

    def run_full_analysis(self):
        """執行完整分析"""
        print("🚀 開始庫存分析...")

        if not self.load_data():
            return False

        # 執行各項分析
        self.basic_statistics()
        self.warehouse_analysis()
        self.trend_analysis()
        self.product_analysis()
        self.anomaly_detection()
        self.generate_visualizations()

        # 生成報告
        report_file = self.generate_report()

        print("\n" + "="*50)
        print("🎉 分析完成！")
        print("="*50)
        print(f"📋 報告檔案: {report_file}")
        print("📊 圖表目錄: charts/")
        print("\n建議每日執行此分析以監控庫存變化。")

        return True


def main():
    """主程式"""
    analyzer = InventoryAnalyzer()
    analyzer.run_full_analysis()


if __name__ == "__main__":
    main()
