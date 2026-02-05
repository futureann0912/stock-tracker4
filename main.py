#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
台灣股市即時追蹤器 - Windows 桌面應用程式 (Tkinter 版本)
Taiwan Stock Market Real-time Tracker - Desktop Application (Tkinter Version)
"""

import sys
import json
import threading
import subprocess
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# 自動安裝 requests 依賴
def install_requests():
    """檢查並安裝 requests 依賴"""
    try:
        import requests
    except ImportError:
        print("正在安裝 requests 依賴...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "requests", "-q"])
            print("requests 安裝完成")
        except subprocess.CalledProcessError:
            messagebox.showerror("錯誤", "無法安裝 requests 依賴")
            sys.exit(1)

install_requests()

import requests


class StockTrackerApp:
    """台灣股市追蹤器主應用程式"""
    
    def __init__(self, root):
        self.root = root
        self.root.title('台灣股市即時追蹤器')
        self.root.geometry('1000x600')
        
        self.stocks = {}
        self.config_file = Path.home() / '.stock_tracker_config.json'
        self.auto_refresh_active = False
        self.refresh_interval = 30000  # 毫秒
        
        self.init_ui()
        self.load_stocks()
    
    def init_ui(self):
        """初始化用戶界面"""
        # 主框架
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 標題
        title_label = ttk.Label(main_frame, text='台灣股市即時追蹤器', font=('Arial', 14, 'bold'))
        title_label.pack(pady=(0, 10))
        
        # 控制面板
        control_frame = ttk.Frame(main_frame)
        control_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 股票代碼輸入
        ttk.Label(control_frame, text='股票代碼:').pack(side=tk.LEFT, padx=(0, 5))
        self.stock_input = ttk.Entry(control_frame, width=10)
        self.stock_input.pack(side=tk.LEFT, padx=(0, 5))
        self.stock_input.bind('<Return>', lambda e: self.add_stock())
        
        # 新增按鈕
        ttk.Button(control_frame, text='新增', command=self.add_stock).pack(side=tk.LEFT, padx=2)
        
        # 刷新按鈕
        ttk.Button(control_frame, text='刷新', command=self.refresh_all_stocks).pack(side=tk.LEFT, padx=2)
        
        # 自動刷新
        ttk.Label(control_frame, text='自動刷新:').pack(side=tk.LEFT, padx=(20, 5))
        self.auto_refresh_var = tk.BooleanVar()
        self.auto_refresh_check = ttk.Checkbutton(
            control_frame, 
            variable=self.auto_refresh_var,
            command=self.toggle_auto_refresh
        )
        self.auto_refresh_check.pack(side=tk.LEFT, padx=(0, 10))
        
        # 刷新間隔
        ttk.Label(control_frame, text='間隔(秒):').pack(side=tk.LEFT, padx=(0, 5))
        self.interval_var = tk.StringVar(value='30')
        interval_spinbox = ttk.Spinbox(
            control_frame,
            from_=5,
            to=300,
            textvariable=self.interval_var,
            width=5,
            command=self.update_refresh_interval
        )
        interval_spinbox.pack(side=tk.LEFT, padx=(0, 20))
        
        # 匯出按鈕
        ttk.Button(control_frame, text='匯出', command=self.export_data).pack(side=tk.LEFT, padx=2)
        
        # 表格框架
        table_frame = ttk.Frame(main_frame)
        table_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # 創建表格
        columns = ('代碼', '名稱', '現價', '漲跌', '開盤', '最高', '最低', '成交量', '更新時間')
        self.tree = ttk.Treeview(table_frame, columns=columns, height=15, show='headings')
        
        # 定義列
        for col in columns:
            self.tree.column(col, width=100)
            self.tree.heading(col, text=col)
        
        # 滾動條
        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 底部按鈕框架
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.pack(fill=tk.X)
        
        ttk.Button(bottom_frame, text='刪除選中', command=self.delete_selected).pack(side=tk.LEFT, padx=2)
        ttk.Button(bottom_frame, text='清除全部', command=self.clear_all).pack(side=tk.LEFT, padx=2)
        
        status_label = ttk.Label(bottom_frame, text='提示: 輸入股票代碼後按 Enter 或點擊「新增」按鈕')
        status_label.pack(side=tk.RIGHT, padx=2)
    
    def add_stock(self):
        """添加股票到追蹤列表"""
        code = self.stock_input.get().strip().upper()
        if not code:
            messagebox.showwarning('警告', '請輸入股票代碼')
            return
        
        if not code.isdigit() or len(code) != 4:
            messagebox.showwarning('警告', '股票代碼必須是 4 位數字')
            return
        
        if code in self.stocks:
            messagebox.showwarning('警告', f'股票 {code} 已在追蹤列表中')
            return
        
        # 在新線程中獲取股票數據
        thread = threading.Thread(target=self.fetch_stock, args=(code,))
        thread.daemon = True
        thread.start()
        
        self.stock_input.delete(0, tk.END)
    
    def fetch_stock(self, code):
        """從 TWSE API 獲取股票數據"""
        try:
            date_str = datetime.now().strftime('%Y%m%d')
            url = f"https://www.twse.com.tw/exchangeReport/STOCK_DAY?response=json&date={date_str}&stockNo={code}"
            
            response = requests.get(url, timeout=10)
            response.raise_for_status()
            data = response.json()
            
            if data.get('data') and len(data['data']) > 0:
                latest = data['data'][-1]
                stock_info = {
                    'code': code,
                    'name': data.get('name', code),
                    'price': float(latest[6]),
                    'change': float(latest[7]),
                    'volume': int(latest[1].replace(',', '')),
                    'open': float(latest[3]),
                    'high': float(latest[4]),
                    'low': float(latest[5]),
                    'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                }
                self.stocks[code] = stock_info
                self.update_table()
                self.save_stocks()
            else:
                messagebox.showerror('錯誤', f"找不到股票代碼: {code}")
        except requests.exceptions.Timeout:
            messagebox.showerror('錯誤', f"連線超時: {code}")
        except requests.exceptions.ConnectionError:
            messagebox.showerror('錯誤', "網路連線失敗，請檢查網路設定")
        except Exception as e:
            messagebox.showerror('錯誤', f"錯誤: {str(e)}")
    
    def refresh_all_stocks(self):
        """刷新所有股票數據"""
        if not self.stocks:
            return
        
        for code in self.stocks.keys():
            thread = threading.Thread(target=self.fetch_stock, args=(code,))
            thread.daemon = True
            thread.start()
    
    def toggle_auto_refresh(self):
        """切換自動刷新"""
        if self.auto_refresh_var.get():
            self.schedule_refresh()
        else:
            self.auto_refresh_active = False
    
    def update_refresh_interval(self):
        """更新刷新間隔"""
        try:
            self.refresh_interval = int(self.interval_var.get()) * 1000
            if self.auto_refresh_var.get():
                self.schedule_refresh()
        except ValueError:
            pass
    
    def schedule_refresh(self):
        """安排自動刷新"""
        if self.auto_refresh_var.get():
            self.refresh_all_stocks()
            self.root.after(self.refresh_interval, self.schedule_refresh)
    
    def update_table(self):
        """更新表格顯示"""
        # 清空表格
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # 添加數據
        for code, stock in self.stocks.items():
            change = stock['change']
            change_str = f"{change:+.2f}"
            
            values = (
                code,
                stock['name'],
                f"{stock['price']:.2f}",
                change_str,
                f"{stock['open']:.2f}",
                f"{stock['high']:.2f}",
                f"{stock['low']:.2f}",
                f"{stock['volume']:,}",
                stock['timestamp']
            )
            
            # 根據漲跌設置顏色
            if change > 0:
                self.tree.insert('', tk.END, values=values, tags=('up',))
            elif change < 0:
                self.tree.insert('', tk.END, values=values, tags=('down',))
            else:
                self.tree.insert('', tk.END, values=values)
        
        # 配置標籤顏色
        self.tree.tag_configure('up', foreground='red')
        self.tree.tag_configure('down', foreground='green')
    
    def delete_selected(self):
        """刪除選中的股票"""
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning('警告', '請先選擇要刪除的股票')
            return
        
        for item in selected:
            values = self.tree.item(item, 'values')
            code = values[0]
            if code in self.stocks:
                del self.stocks[code]
        
        self.update_table()
        self.save_stocks()
    
    def clear_all(self):
        """清除所有股票"""
        if messagebox.askyesno('確認', '確定要清除所有追蹤的股票嗎?'):
            self.stocks.clear()
            self.update_table()
            self.save_stocks()
    
    def export_data(self):
        """匯出股票數據為 CSV"""
        if not self.stocks:
            messagebox.showwarning('警告', '沒有股票數據可匯出')
            return
        
        file_path = filedialog.asksaveasfilename(
            defaultextension='.csv',
            filetypes=[('CSV Files', '*.csv'), ('All Files', '*.*')]
        )
        
        if not file_path:
            return
        
        try:
            with open(file_path, 'w', encoding='utf-8-sig') as f:
                f.write('代碼,名稱,現價,漲跌,開盤,最高,最低,成交量,更新時間\n')
                
                for code, stock in self.stocks.items():
                    f.write(f"{code},{stock['name']},{stock['price']:.2f},{stock['change']:+.2f},"
                            f"{stock['open']:.2f},{stock['high']:.2f},{stock['low']:.2f},"
                            f"{stock['volume']:,},{stock['timestamp']}\n")
            
            messagebox.showinfo('成功', f'數據已匯出到: {file_path}')
        except Exception as e:
            messagebox.showerror('錯誤', f'匯出失敗: {str(e)}')
    
    def save_stocks(self):
        """將股票列表保存到配置文件"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(list(self.stocks.keys()), f)
        except Exception as e:
            print(f'保存配置失敗: {e}')
    
    def load_stocks(self):
        """從配置文件加載股票列表"""
        if self.config_file.exists():
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    codes = json.load(f)
                    for code in codes:
                        thread = threading.Thread(target=self.fetch_stock, args=(code,))
                        thread.daemon = True
                        thread.start()
            except Exception as e:
                print(f'加載配置失敗: {e}')


def main():
    root = tk.Tk()
    app = StockTrackerApp(root)
    root.mainloop()


if __name__ == '__main__':
    main()
