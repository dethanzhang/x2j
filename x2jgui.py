# !/usr/bin/python3
# -*- coding: utf-8 -*-

import os
import tkinter as tk
from tkinter import ttk, messagebox
import x2jcore, x2jutils

version = "3.4"
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

# 获取屏幕尺寸
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

# 创建主容器Frame
main_container = tk.Frame(root)
main_container.pack(pady=10, fill="both", expand=True)

# 创建Canvas和Scrollbar
canvas = tk.Canvas(main_container)
scrollbar = ttk.Scrollbar(main_container, orient="vertical", command=canvas.yview)
data_frame = tk.Frame(canvas)

# 配置Canvas和滚动区域
canvas.configure(yscrollcommand=scrollbar.set)


# 设置窗口的最小尺寸和初始尺寸
def update_window_size(event=None):
    # 计算内容需要的尺寸
    data_frame.update_idletasks()
    content_width = data_frame.winfo_reqwidth() + scrollbar.winfo_reqwidth() + 40

    # 计算所有控件的高度（更精确的计算）
    canvas_height = data_frame.winfo_reqheight()
    # 下面的控件高度估计：勾选框(30) + 功能框架(20*n) + 执行按钮(80) + 边距(40)
    bottom_widgets_height = 30 + len(all_json_folder) * 25 + 150 + 40
    content_height = min(
        canvas_height + bottom_widgets_height, int(screen_height * 0.85)
    )

    # 限制窗口的最大宽度为屏幕宽度的80%
    max_width = int(screen_width * 0.8)
    window_width = min(content_width, max_width)

    # 设置窗口的最小尺寸
    root.minsize(int(screen_width * 0.4), 500)

    # 设置窗口的初始尺寸
    root.geometry(f"{window_width}x{int(content_height)}")

    # 设置Canvas的宽度
    canvas.configure(width=window_width - scrollbar.winfo_reqwidth() - 10)


# 绑定窗口大小更新函数
data_frame.bind("<Configure>", update_window_size)

# 创建变量来存储每个勾选框的状态
item_vars = []

target_string = "文本"  # 文件名包含该字符的会另外显示在特殊区域

# 统计普通配置表和文本配置表的数量
normal_files = [x for x in all_xlsx_path if target_string not in x]
text_files = [x for x in all_xlsx_path if target_string in x]
normal_count = len(normal_files)
text_count = len(text_files)


# 根据文件数量确定列数
def get_column_count(file_count):
    if file_count <= 20:
        return 2
    elif file_count <= 30:
        return 3
    else:
        return 4


normal_columns = get_column_count(normal_count)
text_columns = get_column_count(text_count)


# 计算每列应该显示的行数
def get_rows_per_column(total_items, num_columns):
    base_rows = total_items // num_columns
    extra = total_items % num_columns
    return [base_rows + (1 if i < extra else 0) for i in range(num_columns)]


# 获取每列的行数
normal_rows_per_column = get_rows_per_column(normal_count, normal_columns)
text_rows_per_column = get_rows_per_column(text_count, text_columns)

# 在普通文件区域上方添加标题
normal_label = tk.Label(
    data_frame,
    text=f"--- 普通配置表 ({normal_count}) ---",
    fg="green",
    font=("Arial", 12, "bold"),
)
normal_label.grid(
    row=0, column=0, columnspan=max(normal_columns, text_columns), pady=7, sticky="w"
)

# 添加普通配置表
current_row = 1
current_col = 0
files_in_col = 0

for item in normal_files:
    check_var = tk.BooleanVar()
    item_vars.append(check_var)

    tk.Checkbutton(data_frame, text=item, variable=check_var).grid(
        row=current_row, column=current_col, padx=5, pady=2, sticky=tk.W
    )
    files_in_col += 1
    current_row += 1

    if files_in_col >= normal_rows_per_column[current_col]:
        current_col += 1
        current_row = 1
        files_in_col = 0

# 计算文本配置表的起始行
text_start_row = max([rows for rows in normal_rows_per_column]) + 3

# 添加文本配置表标题
separator_label = tk.Label(
    data_frame,
    text=f"--- {target_string}配置表 ({text_count}) ---",
    fg="blue",
    font=("Arial", 12, "bold"),
)
separator_label.grid(
    row=text_start_row,
    column=0,
    columnspan=max(normal_columns, text_columns),
    pady=7,
    sticky="w",
)

# 添加文本配置表
current_row = text_start_row + 1
current_col = 0
files_in_col = 0

for item in text_files:
    check_var = tk.BooleanVar()
    item_vars.append(check_var)

    tk.Checkbutton(data_frame, text=item, variable=check_var).grid(
        row=current_row, column=current_col, padx=5, pady=2, sticky=tk.W
    )
    files_in_col += 1
    current_row += 1

    if files_in_col >= text_rows_per_column[current_col]:
        current_col += 1
        current_row = text_start_row + 1
        files_in_col = 0

# 设置Canvas和Scrollbar
canvas.create_window((0, 0), window=data_frame, anchor="nw")


# 配置Canvas的滚动区域和更新窗口大小
def on_frame_configure(event):
    canvas.configure(scrollregion=canvas.bbox("all"))
    update_window_size()


data_frame.bind("<Configure>", on_frame_configure)

# 放置Canvas和Scrollbar
canvas.pack(side="left", fill="both", expand=True)
scrollbar.pack(side="right", fill="y")


# 计算实际使用的最后一行
def get_last_row():
    # 获取 data_frame 中所有子控件的最大行号
    if data_frame.winfo_children():
        max_row = 0
        for widget in data_frame.winfo_children():
            grid_info = widget.grid_info()
            if grid_info:
                row = grid_info.get("row", 0)
                if row > max_row:
                    max_row = row
        return max_row
    return 30


# 在特殊文件区域之前添加一条分隔线（动态计算行号）
separator = ttk.Separator(data_frame, orient="horizontal")
last_row = get_last_row() + 1
separator.grid(row=last_row, column=0, columnspan=5, sticky="ew", pady=10)

# 创建底部框架用于管理功能选项和执行按钮
bottom_frame = tk.Frame(root)
bottom_frame.pack(fill="both", expand=False, padx=10, pady=10)

# 功能勾选框 - 是否开启自动移动到对应目录
enable_var = tk.IntVar(value=0)  # 用来存储 "开启选择" 勾选框的状态，默认未勾选
enable_checkbox = tk.Checkbutton(
    bottom_frame,
    text="是否自动导出到指定json目录(文件会被直接覆盖)",
    variable=enable_var,
    command=toggle_selection,
)
enable_checkbox.pack(pady=5, anchor="center")

# 创建一个Frame用于放置功能选项
feature_frame = tk.Frame(bottom_frame)
feature_frame.pack(fill="both", expand=False, pady=5)

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
    radio.pack(anchor="center", pady=2)
    radio_buttons.append(radio)

# 创建一个按钮用于执行脚本
execute_button = tk.Button(
    bottom_frame, text="开始导表", command=perform_action, width=15, height=3
)
execute_button.pack(side=tk.BOTTOM, pady=10, anchor="center")

# 运行主循环
root.mainloop()
