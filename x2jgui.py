# !/usr/bin/python3
# -*- coding: utf-8 -*-
#需要安装xlrd 1.2.0版本。#python3 -m pip install xlrd==1.2.0

import os
import json
import tkinter as tk
from tkinter import ttk, messagebox
import xlrd

xlrd.book.encoding='gbk'
CONST_TITLE_POS = 1
CONST_TYPE_POS = 2
CONST_SUBTYPE_POS = 3
CONST_CONTENT_BEGIN = 5
CONST_EXCEL_PATH = "配置表"
CONST_OUTPUT_PATH = 'output_temp'
CONST_SAVE_PATH = 'json'
CONST_FILES_LIST = []
CONST_ERROR_MSG = []
CONST_ERROR_CNT = 0
CONST_CURRENT_SHEET = ''

#mac系统打包成.app需要特殊处理当前路径
now = os.getcwd()
if '.app' in now:
    os.chdir(os.path.split(now.split('.app')[0])[0])

def file_extension(path): 
    return os.path.splitext(path)[1]

def clearTempFiles(filePath):
    if os.path.exists(CONST_OUTPUT_PATH):
        #如果存在的话清空里面所有json/txt文件
        list_files = os.listdir(CONST_OUTPUT_PATH)
        for i in range(0,len(list_files)):
            file = list_files[i]
            subfilePath = os.path.join(filePath,file)
            if os.path.isfile(subfilePath) and (file_extension(file) == '.json' or file_extension(file) == '.txt'):
                os.remove(subfilePath)
    else:
        os.makedirs(filePath)

def getAllFolder(prefix): #遍历路径下所有特定前缀的目录
    return [x for x in os.listdir() if x.startswith(prefix) and os.path.isdir(x)]

def pathFileList(filePath):
    listDir = sorted(os.listdir(filePath))
    filenames = []
    for x in listDir:
        if (not x.startswith('~$')) and (file_extension(x) == '.xlsx' or file_extension(x) == '.xls'):
            filenames.append(x)
    return filenames

# 获得所有json目录路径
all_json_path = getAllFolder(CONST_SAVE_PATH) 

#获得所有excel文件路径
all_xlsx_path = []
xlsx_folder_list = getAllFolder(CONST_EXCEL_PATH)
for a in xlsx_folder_list:
    all_xlsx_path += [os.path.join(a,b) for b in pathFileList(a)]

#清理temp目录
clearTempFiles(CONST_OUTPUT_PATH)

def autoMove(dir2):
    list_files = os.listdir(CONST_OUTPUT_PATH)
    for x in list_files:
        try:
            if os.path.exists(os.path.join(dir2,x)):
                os.remove(os.path.join(dir2,x))
            os.rename(os.path.join(CONST_OUTPUT_PATH,x),os.path.join(dir2,x))
        except Exception as e:
            CONST_ERROR_MSG.append(str(e))

def readExcel(filename):
    workbook = xlrd.open_workbook(filename)
    sheet_names= workbook.sheet_names()
    for sheet in sheet_names:
        if sheet == '':
            continue
        outputs = sheet.split('|')
        if len(outputs) <= 1:
            continue
        sheet_name = outputs[0]
        singleOutput = 0
        if sheet_name.startswith('#'): #基础单层级json
            singleOutput = 1
            sheet_name = sheet_name[1:]
        if sheet_name.startswith('^'): #两层级json
            singleOutput = 2
            sheet_name = sheet_name[1:]
        if sheet_name.startswith('$'): #基础文本表
            singleOutput = 3
            sheet_name = sheet_name[1:]
        # if sheet_name.startswith('%'):
        #     singleOutput = 4
        #     sheet_name = sheet_name[1:]

        sheet_data = workbook.sheet_by_name(sheet)
        titles = sheet_data.row_values(CONST_TITLE_POS) #拿到表头
        types = sheet_data.row_values(CONST_TYPE_POS) #拿到类型
        subTypes = sheet_data.row_values(CONST_SUBTYPE_POS) #拿到子类型备注

        find = -1
        findLevel = -1
        findGroup = -1
        findArrayGroup = -1

        for i in range(0,len(titles)):
            title = titles[i]
            if ' ' in title:
                CONST_ERROR_MSG.append('%s表中的%s字段包含有空格, 请检查'%(sheet_name,title))
                global CONST_ERROR_CNT
                CONST_ERROR_CNT += 1
                return None
            if title.startswith('*'):
                titles[i] = title[1:]
                find = i
            if title.startswith('^'):
                titles[i] = title[1:]
                findLevel = i
            # if title.startswith('%'):
            #     titles[i] = title[1:]
            #     findGroup = i
            # if title.startswith('[]'):
            #     titles[i] = title[2:]
            #     findArrayGroup = i

        global CONST_CURRENT_SHEET
        CONST_CURRENT_SHEET = sheet
        #单独输出sheet为整个json文件 并以id为key包含每行内容为对象进行序列化
        if singleOutput==1:
            if findGroup != -1:
                if find != -1:
                    # table = readExcelByGroup(findGroup,find,titles,types,subTypes,sheet_data)
                    # writeJsonFile(sheet_name+".json",table)
                    pass
            elif find != -1:
                # table = readExcelByKey(find,titles,types,subTypes,sheet_data)
                # writeJsonFile(sheet_name+".json",table)
                pass
            else :
                #都没有找到 为无key文件
                table = readExcelNoKey(titles,types,subTypes,sheet_data)
                writeJsonFile(sheet_name+".json",table)
            continue

        if singleOutput == 2:
            if findLevel != -1:
                    table = readExcelByLevel(findLevel,titles,types,subTypes,sheet_data)
                    writeJsonFile(sheet_name+'.json',table)
            continue

        if singleOutput == 3:
            try:
                readLocalizationExcel(titles,sheet_data,workbook.sheet_by_name('异常字符集'))
                print("已读取错误字符集")
            except:
                readLocalizationExcel(titles,sheet_data)
            continue

    return None

# 将单个sheet输出为一个数组, 一行数据为一个字典对象
def readExcelNoKey(titles,types,subTypes,sheet_data):
    table = []
    row_num = sheet_data.nrows
    for i in range(CONST_CONTENT_BEGIN,row_num):
        content = sheet_data.row_values(i)
        if content[0] == '':
            continue
        line = {}
        for j in range(0,len(titles)):
            if titles[j].startswith('#'):
                continue
            if types[j] != '':
                line[titles[j]] = getValueByType(content[j],types[j],subTypes[j],i,j)
        table.append(line)
    return table

# 将单个sheet输出为一个数组, 根据标记位置, 将多行组成一个数组. 只支持2个level,即[{ 主键 [{从属内容},] },]
def readExcelByLevel(findLevel,titles,types,subTypes,sheet_data):
    table = []
    row_num = sheet_data.nrows
    levelGroup = []

    for i in range(CONST_CONTENT_BEGIN,row_num):
        content = sheet_data.row_values(i)
        line = {}

        if content[0] != '':#首列不为空,表示有新增主键
            try:
                table.append(alevel)
            except:
                pass
            alevel = {}
            for j in range(0,findLevel):#创建主键内容
                if titles[j].startswith('#'):
                    continue
                alevel[titles[j]] = getValueByType(content[j],types[j],subTypes[j])
            alevel[titles[findLevel]] = []

        for j in range(findLevel+1, len(titles)):#将每行的从属内容合并
            if titles[j].startswith('#'):
                continue
            line[titles[j]] = getValueByType(content[j],types[j],subTypes[j],i,j)
        alevel[titles[findLevel]].append(line)
    table.append(alevel)
    return table

# **文本表专用** 读取单个sheet, 以首列为字典key, 后续每一列作为不同dict的value, 并以每一列的id作为json名输出
def readLocalizationExcel(titles,sheet_data,char_data=None):
    keys = sheet_data.col_values(0)[CONST_CONTENT_BEGIN:]
    col_num = sheet_data.ncols
    for i in range(1,col_num):
        if titles[i].startswith('#'):
            continue
        content = sheet_data.col_values(i)[CONST_CONTENT_BEGIN:]
        if char_data:
            list_badChar = char_data.col_values(0)[1:]
            # list_badChar_unicode = char_data.col_values(1)[1:]
            list_goodChar = char_data.col_values(2)[1:]
            # list_goodChar_unicode = char_data.col_values(3)[1:]
            content = fixBadChar(content,list_badChar,list_goodChar)
        table = dict(zip(keys,content))
        writeJsonFile(titles[i]+'.json',table)


def fixBadChar(content,list_badChar,list_goodChar):
    for i,x in enumerate(list_badChar):
        for idx,c in enumerate(content):
            if x in str(c):
                content[idx] = c.replace(x,list_goodChar[i])
                print("检测到错误")
    return content


def writeJsonFile(filename,data):
    data_str = json.dumps(data, sort_keys=False, indent=4, ensure_ascii=False)
    filePath = os.path.join(CONST_OUTPUT_PATH,filename)
    if os.path.exists(filePath):
        os.remove(filePath)
    with open(filePath,'w',encoding='utf8') as f:
        f.write(data_str)
    if os.name == 'posix': #非windows系统执行, 将LF转换为CRLF
        with open(filePath,'rb') as f:
            content = f.read()
        content = content.replace(b'\n',b'\r\n')
        with open(filePath,'wb') as f:
            f.write(content)



def autoValue(value):
    try:
        value = float(value)
    except ValueError:
        return value
    value = '{:.12g}'.format(value)
    try:#是否整数
        return int(value)
    except ValueError:
        pass
    return float(value)

def getValueByType(value,type1,subType=None,row=None,column=None):
    try:
        if value == 'null':
            if type1 == 'matrix' or type1 == 'array' or type1 == 'array-str' or type1 == 'json' or type1 == 'json-str':
                return []
            if type1 == 'int' or type1 == 'float':
                return 0
            return ''
        if type1 == "int":
            if value == '':
                return 0
            return int(value)
        if type1 == "float":
            return float(value)
        if type1 == "bool":
            if value == "true" or autoValue(value) == 1 or value == "True" or value == "TRUE":
                return True
            return False

        if type1 == "array-str":
            if '\n' in str(value):
                value=value.replace('\n',',')
            if ',' not in str(value):
                return [str(autoValue(value))]
            listsValue = [str(autoValue(i)) for i in value.split(',')]
            return listsValue

        if type1 == "array":
            if '\n' in str(value):
                value=value.replace('\n',',')
            if value == "":
                return []
            if ',' not in str(value):
                value = autoValue(value)
                return [value]
            listsValue = [autoValue(i) for i in value.split(',')]
            return listsValue

        if type1 == 'matrix':
            if value == 0:
                return []
            if '\n' in str(value):
                value=value.replace('\n',',')
            if '|' not in str(value):
                value = getValueByType(value,'array',None)
                return [value]

            matrixList = [x for x in value.split('|')]
            for i in range(len(matrixList)):
                matrixList[i] = getValueByType(matrixList[i],'array',None)
            return matrixList

        if type1 == "json-str":
            if subType:
                maps = subType.split(',')
                strs = value.split(';')
                ret = []
                if value == "":
                    return ret
                for i in range(0,len(strs)):
                    dic = {}
                    contents = strs[i].split(',')
                    for j in range(0,len(maps)):
                        dic[maps[j]] = getValueByType(contents[j],"str",None)
                    ret.append(dic)
                return ret
            else:
                return json.loads(value)

        if type1 == "json":
            if value == '':
                return []
            if subType:
                maps = subType.split(',')
                strs = value.split(';')
                ret = []
                for i in range(0,len(strs)):
                    dic = {}
                    contents = strs[i].split(',')
                    for j in range(0,len(maps)):
                        dic[maps[j]] = autoValue(contents[j])
                    ret.append(dic)
                return ret
            else:
                return json.loads(value)


        if type1 == "str" or type1 == "string":
            if type(value).__name__ == 'str':
                return value
            value = autoValue(value)
            return str(value)

    except:
        CONST_ERROR_MSG.append("第【%d】行 第【%d】列单元格填写错误"%(row+1,column+1))
        global CONST_ERROR_CNT
        CONST_ERROR_CNT += 1
        return None


###########################################---GUI部分---###########################################

def toggle_selection():
    # 根据 "开启选择" 勾选框的状态，批量启用或禁用所有单选按钮
    new_state = tk.NORMAL if enable_var.get() == 1 else tk.DISABLED
    for radio in radio_buttons:
        radio.config(state=new_state)
    if new_state == tk.DISABLED:
        feature_var.set(0)  # 重置单选按钮的选项

def perform_action():#执行脚本
    #清理temp目录
    clearTempFiles(CONST_OUTPUT_PATH)
    # 获取勾选的表格路径
    selected_items = [item_var.get() for item_var in item_vars]
    try:
        # 执行导表
        for idx, flag in enumerate(selected_items):
            if flag == True:
                print('正在导出 >> %s'%all_xlsx_path[idx])
                readExcel(all_xlsx_path[idx])
                if len(CONST_ERROR_MSG) > 0:
                    msg = "《%s》-【%s】\n-------------------\n"%(all_xlsx_path[idx],CONST_CURRENT_SHEET) + "\n\n".join(CONST_ERROR_MSG)
                    messagebox.showwarning("错误", msg)
                    CONST_ERROR_MSG.clear()
                if feature_var.get() > 0:
                    dir2 = all_json_path[feature_var.get()-1]
                    #执行移动
                    autoMove(dir2)
                    if len(CONST_ERROR_MSG) > 0:
                        msg = "《%s》\n-------------------\n"%all_xlsx_path[idx] + "\n\n".join(CONST_ERROR_MSG)
                        messagebox.showwarning("错误", msg)
                        CONST_ERROR_MSG.clear()
    except Exception as e:
        messagebox.showerror("错误", str(e))
    
    global CONST_ERROR_CNT
    if CONST_ERROR_CNT == 0:
        # 显示一个消息框，告知用户操作已完成
        messagebox.showinfo("操作完成", "导表完毕")
    else:
        messagebox.showinfo("操作完成", "出现%s个错误, 建议检查后重新执行"%CONST_ERROR_CNT)
    CONST_ERROR_CNT = 0

# 创建主窗口
root = tk.Tk()
root.title("导表工具")

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
for idx, feature in enumerate(all_json_path):
    radio = tk.Radiobutton(feature_frame, text='自动导出到>>%s'%feature, variable=feature_var, value=idx+1, state=tk.DISABLED)
    radio.pack(anchor='w')
    radio_buttons.append(radio)

# 创建一个按钮用于执行脚本
execute_button = tk.Button(root, text="开始导表", command=perform_action, width=30, height=2)
execute_button.pack(side=tk.BOTTOM, pady=10)


# 运行主循环
root.mainloop()
