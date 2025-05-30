# -*- coding: utf-8 -*-
# @Author : 小红牛
# 微信公众号：WdPython
import tkinter as tk
from tkcalendar import DateEntry

def update_date(event=None):
    """当日期改变时自动更新显示"""
    selected_date = cal.get_date()
    result_label.config(text=f"当前选择: {selected_date.strftime('%Y-%m-%d')}")

# 创建主窗口
root = tk.Tk()
root.title("tkcalendar日期组件显示")
root.geometry("300x200")

# 创建日期选择框并绑定事件
cal = DateEntry(root,
               width=12,
               background='darkblue',
               foreground='white',
               borderwidth=2,
               date_pattern='yyyy-mm-dd')
cal.pack(pady=20)

# 绑定日期选择事件
cal.bind("<<DateEntrySelected>>", update_date)

# 初始化显示
result_label = tk.Label(root, text="请选择日期", font=('Arial', 10))
result_label.pack(pady=10)

# 设置初始值（可选）
cal.set_date("2025-02-20")  # 设置默认日期
update_date()  # 初始化显示

root.mainloop()
