# !/usr/bin/python3
# -*- coding: utf-8 -*-

import os
import json
import tkinter as tk
from tkinter import ttk, messagebox
import x2jcore, x2jutils

version = '2.2'
ax = x2jcore.x2jcore()
x2jutils.checkChdir()

def toggle_selection():
    # 根据 "开启选择" 勾选框的状态，批量启用或禁用所有单选按钮
    new_state = tk.NORMAL if enable_var.get() == 1 else tk.DISABLED
    for radio in radio_buttons:
        radio.config(state=new_state)
    if new_state == tk.DISABLED:
        feature_var.set(0)  # 重置单选按钮的选项

def print_error(xlsx_path):
    global ax
    for sheetname, errormsgs in ax.error_msg.items():
        if len(errormsgs) > 0:
            msg = "《%s》-【%s】\n-------------------\n"%(xlsx_path,sheetname) + "\n\n".join(errormsgs)
            messagebox.showwarning("错误", msg)



def perform_action():#执行脚本
    global ax
    #清理temp目录
    x2jutils.clearTempFiles(ax.output_path)
    # 获取勾选的表格路径
    selected_items = [item_var.get() for item_var in item_vars]
    try:
        # 执行导表
        for idx, flag in enumerate(selected_items):
            if flag == True:
                print('正在导出 >> %s'%all_xlsx_path[idx])
                ax.start(all_xlsx_path[idx])
                if ax.error_cnt > 0:
                    print_error(all_xlsx_path[idx])

                if feature_var.get() > 0:
                    dir2 = all_json_folder[feature_var.get()-1]
                    #执行移动
                    r = x2jutils.autoMove(ax.output_path,dir2) #成功返回None，失败返回错误信息
                    if r:
                        messagebox.showwarning("错误", "移动文件失败 >> %s"%r)
    except Exception as e:
        messagebox.showerror("工具运行错误", str(e))
        return None
    
    if ax.error_cnt == 0:
        # 显示一个消息框，告知用户操作已完成
        messagebox.showinfo("操作完成", "导表完毕")
    else:
        messagebox.showinfo("操作完成", "出现%s个错误, 建议检查后重新执行"%ax.error_cnt)
    ax.error_cnt = 0




# 获得所有json目录路径
all_json_folder = x2jutils.getAllFolders(ax.save_path) 

all_xlsx_path = []
#获得所有excel文件路径
for a in x2jutils.getAllFolders(ax.excel_path):
    all_xlsx_path += [os.path.join(a,b) for b in x2jutils.xlsxFileList(a)]

#清理temp目录
x2jutils.clearTempFiles(ax.output_path)

# 创建主窗口
root = tk.Tk()
root.title("导表工具 v%s"%version)

# 创建一个Frame用于放置数据列表
data_frame = tk.Frame(root)
data_frame.pack(pady=10)

# 创建变量来存储每个勾选框的状态
item_vars = []

# 创建勾选框并添加到列表中
r = 0
c = 0
for item in all_xlsx_path:
    check_var = tk.BooleanVar()
    item_vars.append(check_var)
    # tk.Checkbutton(data_frame, text=item, variable=check_var).pack(anchor='w')
    tk.Checkbutton(data_frame, text=item, variable=check_var).grid(row=r, column=c, padx=5, pady=2, sticky=tk.W)
    r += 1
    if r % 8 == 0:
        c += 1
        r = 0

# 功能勾选框 - 是否开启自动移动到对应目录
enable_var = tk.IntVar(value=0)  # 用来存储 "开启选择" 勾选框的状态，默认未勾选
enable_checkbox = tk.Checkbutton(root, text="是否自动导出到指定json目录(文件会被直接覆盖)", variable=enable_var, command=toggle_selection)
enable_checkbox.pack(pady=2)

# 创建一个Frame用于放置功能选项
feature_frame = tk.Frame(root)
feature_frame.pack(pady=5)

# 创建单选按钮（默认禁用状态）
radio_buttons = []  # 用于存放所有单选按钮
feature_var = tk.IntVar(value=0) #用于记录选中路径的索引 需要减去1

# 创建功能勾选框并添加到列表中
for idx, feature in enumerate(all_json_folder):
    radio = tk.Radiobutton(feature_frame, text='自动导出到>>%s'%feature, variable=feature_var, value=idx+1, state=tk.DISABLED)
    radio.pack(anchor='w')
    radio_buttons.append(radio)

# 创建一个按钮用于执行脚本
execute_button = tk.Button(root, text="开始导表", command=perform_action, width=30, height=2)
execute_button.pack(side=tk.BOTTOM, pady=10)

# 运行主循环
root.mainloop()
