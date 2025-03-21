# !/usr/bin/python3
# -*- coding: utf-8 -*-

import os
import tkinter as tk
from tkinter import ttk, messagebox
import x2jcore, x2jutils

version = "3.1"
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
            msg = "《%s》-【%s】\n-------------------\n" % (
                xlsx_path,
                sheetname,
            ) + "\n\n".join(errormsgs)
            messagebox.showwarning("错误", msg)


def perform_action():  # 执行脚本
    global ax
    # 清理temp目录
    x2jutils.clearTempFiles(ax.output_path)
    # 获取勾选的表格路径
    selected_items = [item_var.get() for item_var in item_vars]
    try:
        # 执行导表
        for idx, flag in enumerate(selected_items):
            if flag == True:
                print("正在导出 >> %s" % all_xlsx_path[idx])
                ax.start(all_xlsx_path[idx])
                if ax.error_cnt > 0:
                    print_error(all_xlsx_path[idx])

                if feature_var.get() > 0:
                    dir2 = all_json_folder[feature_var.get() - 1]
                    # 执行移动
                    r = x2jutils.autoMove(
                        ax.output_path, dir2
                    )  # 成功返回None，失败返回错误信息
                    if r:
                        messagebox.showwarning("错误", "移动文件失败 >> %s" % r)
    except Exception as e:
        messagebox.showerror("工具运行错误", str(e))
        return None

    if ax.error_cnt == 0:
        # 显示一个消息框，告知用户操作已完成
        messagebox.showinfo("操作完成", "导表完毕")
    else:
        messagebox.showinfo(
            "操作完成", "出现%s个错误, 建议检查后重新执行" % ax.error_cnt
        )
    ax.error_cnt = 0


# 获得所有json目录路径
all_json_folder = x2jutils.getAllFolders(ax.save_path)

# 方便本地工程调试
try:
    with open("projectpath.txt", "r") as f:
        projpath = f.readline()
    all_json_folder.append(os.path.join(projpath))
    print("本地额外导出目录已调用 >> %s" % projpath)
except:
    pass

all_xlsx_path = []
# 获得所有excel文件路径
for a in x2jutils.getAllFolders(ax.excel_path):
    all_xlsx_path += [os.path.join(a, b) for b in x2jutils.xlsxFileList(a)]

# 清理temp目录
x2jutils.clearTempFiles(ax.output_path)

# 创建主窗口
root = tk.Tk()
root.title("导表工具 v%s" % version)

# 创建一个Frame用于放置数据列表
data_frame = tk.Frame(root)
# data_frame.pack(pady=10)
data_frame.pack(pady=10, fill="both", expand=True)

# 创建变量来存储每个勾选框的状态
item_vars = []

target_string = "文本"  # 文件名包含该字符的会另外显示在特殊区域
normal_start_row = 1  # 普通文件的起始行
special_start_row = 20  # 特殊文件的起始行
special_r, special_c = special_start_row, 0  # 特殊条件文件的行列位置
normal_r, normal_c = normal_start_row, 0  # 普通文件的起始行列位置


# 在普通文件区域上方添加标题
normal_label = tk.Label(
    data_frame, text="--- 普通配置表 ---", fg="green", font=("Arial", 12, "bold")
)
normal_label.grid(row=0, column=0, columnspan=5, pady=7, sticky="w")

# 普通文件和特殊文件之间添加分隔符
separator_label = tk.Label(
    data_frame,
    text="--- %s配置表 ---" % target_string,
    fg="blue",
    font=("Arial", 12, "bold"),
)
separator_label.grid(
    row=special_start_row - 1, column=0, columnspan=5, pady=7, sticky="w"
)


for item in all_xlsx_path:
    check_var = tk.BooleanVar()
    item_vars.append(check_var)

    if target_string in item:  # 如果文件名包含目标字符串
        tk.Checkbutton(data_frame, text=item, variable=check_var).grid(
            row=special_r, column=special_c, padx=5, pady=2, sticky=tk.W
        )
        special_r += 1  # 下一行
        if (special_r - special_start_row) % 4 == 0:  # 如果需要换列
            special_c += 1
            special_r = special_start_row
    else:  # 普通文件
        tk.Checkbutton(data_frame, text=item, variable=check_var).grid(
            row=normal_r, column=normal_c, padx=5, pady=2, sticky=tk.W
        )
        normal_r += 1  # 下一行
        if normal_r % 12 == 0:  # 如果需要换列
            normal_c += 1
            normal_r = normal_start_row


# 在特殊文件区域之前添加一条分隔线
separator = ttk.Separator(data_frame, orient="horizontal")
separator.grid(row=30, column=0, columnspan=5, sticky="ew", pady=10)

# 功能勾选框 - 是否开启自动移动到对应目录
enable_var = tk.IntVar(value=0)  # 用来存储 "开启选择" 勾选框的状态，默认未勾选
enable_checkbox = tk.Checkbutton(
    root,
    text="是否自动导出到指定json目录(文件会被直接覆盖)",
    variable=enable_var,
    command=toggle_selection,
)
enable_checkbox.pack(pady=2)

# 创建一个Frame用于放置功能选项
feature_frame = tk.Frame(root)
feature_frame.pack(pady=5)

# 创建单选按钮（默认禁用状态）
radio_buttons = []  # 用于存放所有单选按钮
feature_var = tk.IntVar(value=0)  # 用于记录选中路径的索引 需要减去1

# 创建功能勾选框并添加到列表中
for idx, feature in enumerate(all_json_folder):
    radio = tk.Radiobutton(
        feature_frame,
        text="自动导出到>>%s" % feature,
        variable=feature_var,
        value=idx + 1,
        state=tk.DISABLED,
    )
    radio.pack(anchor="w")
    radio_buttons.append(radio)

# 创建一个按钮用于执行脚本
execute_button = tk.Button(
    root, text="开始导表", command=perform_action, width=30, height=2
)
execute_button.pack(side=tk.BOTTOM, pady=10)

# 运行主循环
root.mainloop()
