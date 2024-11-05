# !/usr/bin/python3
# -*- coding: utf-8 -*-
#需要安装xlrd 1.2.0版本。#python3 -m pip install xlrd==1.2.0

import os
import json
import xlrd
import x2jutils

xlrd.book.encoding='gbk'


class x2jcore:
    def __init__(self,filename):
        self.title_pos = 1
        self.type_pos = 2
        self.subtype_pos = 3
        self.content_begin = 5
        self.excel_path = "配置表"
        self.output_path = 'output_temp'
        self.save_path = 'json'
        self.files_list = []
        self.error_msg = []
        self.error_cnt = 0
        self.wb = xlrd.open_workbook(filename)
        self.sheet_names = self.wb.sheet_names()
        self.flag = self.readExcel() # 0=导表成功 1=字段名带空格

    def readExcel(self):
        for sheet in self.sheet_names:
            if sheet == '':
                continue
            outputs = sheet.split('|')
            if len(outputs) <= 1:
                continue
            sheet_name = outputs[0]
            singleOutput = 0
            if sheet_name.startswith('#'):
                singleOutput = 1
                sheet_name = sheet_name[1:]
            if sheet_name.startswith('^'):
                singleOutput = 2
                sheet_name = sheet_name[1:]
            if sheet_name.startswith('$'):
                singleOutput = 3
                sheet_name = sheet_name[1:]
            # if sheet_name.startswith('%'):
            #     singleOutput = 4
            #     sheet_name = sheet_name[1:]

            self.current_sheet = sheet
            self.sheet_data = self.wb.sheet_by_name(sheet)
            self.titles = sheet_data.row_values(self.title_pos) #拿到表头
            self.types = sheet_data.row_values(self.type_pos) #拿到类型
            self.subTypes = sheet_data.row_values(self.subtype_pos) #拿到子类型备注

            find = -1
            findLevel = -1
            findGroup = -1
            findArrayGroup = -1

            for i in range(0,len(self.titles)):
                title = self.titles[i]
                if ' ' in title:
                    self.error_msg.append('%s表中的%s字段包含有空格, 请检查'%(self.current_sheet,title))
                    self.error_cnt += 1
                    return 1
                if title.startswith('*'):
                    self.titles[i] = title[1:]
                    find = i
                if title.startswith('^'):
                    self.titles[i] = title[1:]
                    findLevel = i
                # if title.startswith('%'):
                #     self.titles[i] = title[1:]
                #     findGroup = i
                # if title.startswith('[]'):
                #     self.titles[i] = title[2:]
                #     findArrayGroup = i

            #单独输出sheet为整个json文件 并以id为key包含每行内容为对象进行序列化
            if singleOutput==1:
                if findGroup != -1:
                    if find != -1:
                        # table = readExcelByGroup(findGroup,find,titles,types,subTypes,sheet_data)
                        # x2jutils.writeJsonFile(sheet_name+".json",table)
                        pass
                elif find != -1:
                    # table = readExcelByKey(find,titles,types,subTypes,sheet_data)
                    # x2jutils.writeJsonFile(sheet_name+".json",table)
                    pass
                else :
                    #都没有找到 为无key文件
                    table = self.readExcelNoKey()
                    x2jutils.writeJsonFile(sheet_name+".json",table)
                continue

            if singleOutput == 2:
                if findLevel != -1:
                        table = readExcelByLevel(findLevel,titles,types,subTypes,sheet_data)
                        x2jutils.writeJsonFile(sheet_name+'.json',table)
                continue

            if singleOutput == 3:
                try:
                    readLocalizationExcel(titles,sheet_data,self.wb.sheet_by_name('异常字符集'))
                    print("已读取错误字符集")
                except:
                    readLocalizationExcel(titles,sheet_data)
                continue

        return 0

    # 将单个sheet输出为一个数组, 一行数据为一个字典对象
    def readExcelNoKey(self):
        table = []
        row_num = self.sheet_data.nrows
        for i in range(self.content_begin,row_num):
            content = self.sheet_data.row_values(i)
            if content[0] == '':
                continue
            line = {}
            for j in range(0,len(self.titles)):
                if self.titles[j].startswith('#'):
                    continue
                if self.types[j] != '':
                    line[self.titles[j]] = x2jutils.getValueByType(content[j],self.types[j],self.subTypes[j])
                    if not line[self.titles[j]]:
                        self.error_msg.append("第【%d】行 第【%d】列单元格填写错误"%(i,j))
                        self.error_cnt += 1
            table.append(line)
        return table

# 将单个sheet输出为一个数组, 根据标记位置, 将多行组成一个数组. 只支持2个level,即[{ 主键 [{从属内容},] },]
def readExcelByLevel(self,findLevel,titles,types,subTypes,sheet_data):
    table = []
    row_num = sheet_data.nrows
    levelGroup = []

    for i in range(self.content_begin,row_num):
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
                alevel[titles[j]] = x2jutils.getValueByType(content[j],types[j],subTypes[j])
            alevel[titles[findLevel]] = []

        for j in range(findLevel+1, len(titles)):#将每行的从属内容合并
            if titles[j].startswith('#'):
                continue
            line[titles[j]] = x2jutils.getValueByType(content[j],types[j],subTypes[j],i,j)
        alevel[titles[findLevel]].append(line)
    table.append(alevel)
    return table

# **文本表专用** 读取单个sheet, 以首列为字典key, 后续每一列作为不同dict的value, 并以每一列的id作为json名输出
def readLocalizationExcel(titles,sheet_data,char_data=None):
    keys = sheet_data.col_values(0)[self.content_begin:]
    col_num = sheet_data.ncols
    for i in range(1,col_num):
        if titles[i].startswith('#'):
            continue
        content = sheet_data.col_values(i)[self.content_begin:]
        if char_data:
            list_badChar = char_data.col_values(0)[1:]
            # list_badChar_unicode = char_data.col_values(1)[1:]
            list_goodChar = char_data.col_values(2)[1:]
            # list_goodChar_unicode = char_data.col_values(3)[1:]
            content = fixBadChar(content,list_badChar,list_goodChar)
        table = dict(zip(keys,content))
        x2jutils.writeJsonFile(titles[i]+'.json',table)


def fixBadChar(content,list_badChar,list_goodChar):
    for i,x in enumerate(list_badChar):
        for idx,c in enumerate(content):
            if x in str(c):
                content[idx] = c.replace(x,list_goodChar[i])
                print("检测到错误")
    return content

if __name__ == "__main__":
    # 获得所有json目录路径
    all_json_path = x2jutils.getAllFolders(self.save_path) 

    #获得所有excel文件路径
    all_xlsx_path = []
    xlsx_folder_list = x2jutils.getAllFolders(self.excel_path)
    for a in xlsx_folder_list:
        all_xlsx_path += [os.path.join(a,b) for b in x2jutils.pathFileList(a)]

    #清理temp目录
    x2jutils.clearTempFiles(self.output_path)

    #导表
    for i in range(0,len(all_xlsx_path)):
        print(i+1,'---',all_xlsx_path[i])
    print('请输入要导出的表的序号:')
    selectFile=int(input())
    if selectFile>len(all_xlsx_path) or selectFile<1:
        print('输入错误 请重新执行脚本')
    else:
        readExcel(all_xlsx_path[selectFile-1])