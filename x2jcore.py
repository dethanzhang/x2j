# !/usr/bin/python3
# -*- coding: utf-8 -*-
#需要安装xlrd 1.2.0版本。#python3 -m pip install xlrd==1.2.0

import os
import json
import xlrd
import x2jutils

xlrd.book.encoding='gbk'


class x2jcore:
    def __init__(self):
        self.title_pos = 1
        self.type_pos = 2
        self.subtype_pos = 3
        self.content_begin = 5
        self.excel_path = "配置表"
        self.output_path = 'output_temp'
        self.save_path = 'json'

    def storeErrorMsg(self,i,j):
        quotient, remainder = divmod(j, 26)
        #0~25 对应A~Z, 26~51 对应AA~AZ, 52~77 对应BA~BZ
        if quotient == 0:
            col_name = chr(remainder + 65)
        else:
            col_name = chr(quotient + 64) + chr(remainder + 65)
        self.error_msg[self.current_sheet].append("行【%s】 【%s】列 填写错误"%(i+1,col_name))
        self.error_cnt += 1

    def start(self,filename): #返回0=导表成功 1=字段名带空格
        self.error_msg = {} #格式为{sheetname1: [错误信息1,...],sheetname2: [错误信息1,...]}
        self.error_cnt = 0
        self.wb = xlrd.open_workbook(filename)
        self.sheet_names = self.wb.sheet_names()
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
            if sheet_name.startswith('%'):
                singleOutput = 4
                sheet_name = sheet_name[1:]

            self.current_sheet = sheet
            self.sheet_data = self.wb.sheet_by_name(sheet)
            self.titles = self.sheet_data.row_values(self.title_pos) #拿到表头
            self.types = self.sheet_data.row_values(self.type_pos) #拿到类型
            self.subTypes = self.sheet_data.row_values(self.subtype_pos) #拿到子类型备注

            find = -1
            findGroup = -1
            # findArrayGroup = -1

            self.error_msg[sheet] = []
            print("导出子表格-----",sheet_name)
            for i in range(0,len(self.titles)):
                title = self.titles[i]
                if ' ' in title:
                    self.error_msg[self.current_sheet].append('%s表中的%s字段包含有空格, 请检查'%(self.current_sheet,title))
                    self.error_cnt += 1
                    return 1
                if title.startswith('*'):
                    self.titles[i] = title[1:]
                    find = i
                if title.startswith('^'):
                    self.titles[i] = title[1:]
                    findGroup = i
                # if title.startswith('%'):
                #     self.titles[i] = title[1:]
                #     findMarks = i
                # if title.startswith('[]'):
                #     self.titles[i] = title[2:]
                #     findArrayGroup = i

            #单独输出sheet为整个json文件 并以id为key包含每行内容为对象进行序列化
            if singleOutput==1:
                if findGroup != -1:
                    if find != -1:
                        table = self.readExcelByGroup(findGroup,find)
                        x2jutils.writeJsonFile(os.path.join(self.output_path,sheet_name+'.json'),table)
                        pass
                elif find != -1:
                    table = self.readExcelByKey(find)
                    x2jutils.writeJsonFile(os.path.join(self.output_path,sheet_name+'.json'),table)
                    pass
                else :
                    #都没有找到 为无key文件
                    table = self.readExcelNoKey()
                    x2jutils.writeJsonFile(os.path.join(self.output_path,sheet_name+'.json'),table)
                continue

            if singleOutput == 2:
                if findGroup != -1:
                        table = self.readExcelWithGroup(findGroup)
                        x2jutils.writeJsonFile(os.path.join(self.output_path,sheet_name+'.json'),table)
                continue

            if singleOutput == 3:
                try:
                    self.readLocalizationExcel(self.wb.sheet_by_name('异常字符集'))
                    print("已读取错误字符集")
                except:
                    self.readLocalizationExcel()
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
                    if line[self.titles[j]] == None:
                        print("第【%s】行 第【%s】列单元格填写错误: %s"%(i+1,j+1,x2jutils.getValueByType(content[j],self.types[j],self.subTypes[j],debug=True)))
                        self.storeErrorMsg(i,j)
            table.append(line)
        return table

    # 将单个sheet输出为一个数组, 根据标记位置, 将多行组成一个数组. 只支持2个level,即[{ 主键 [{从属内容},] },]
    def readExcelWithGroup(self,findGroup):
        table = []
        row_num = self.sheet_data.nrows

        for i in range(self.content_begin,row_num):
            content = self.sheet_data.row_values(i)
            line = {}

            if content[0] != '':#首列不为空,表示有新增主键
                try: #避免首行报错
                    table.append(alevel)
                except:
                    pass
                alevel = {}
                for j in range(0,findGroup):#创建主键内容
                    if self.titles[j].startswith('#'):
                        continue
                    alevel[self.titles[j]] = x2jutils.getValueByType(content[j],self.types[j],self.subTypes[j])
                    if alevel[self.titles[j]] == None: #处理报错
                        print(x2jutils.getValueByType(content[j],self.types[j],self.subTypes[j],debug=True))
                        self.storeErrorMsg(i,j)
                alevel[self.titles[findGroup]] = []

            for j in range(findGroup+1, len(self.titles)):#将每行的从属内容合并
                if self.titles[j].startswith('#'):
                    continue
                line[self.titles[j]] = x2jutils.getValueByType(content[j],self.types[j],self.subTypes[j])
                if line[self.titles[j]] == None: #处理报错
                    print(x2jutils.getValueByType(content[j],self.types[j],self.subTypes[j],debug=True))
                    self.storeErrorMsg(i,j)
            alevel[self.titles[findGroup]].append(line)
        table.append(alevel)
        return table

    # 将单个sheet输出为一个字典, 根据标记位置的列作为字典key. 每一行的其它内容再嵌套成一个字典. {id1:{title1:v1,title2:v2...},...}
    def readExcelByKey(self,find):
        table = {}
        main_col = self.sheet_data.col_values(find)
        main_type = self.types[find]
        for i in range(self.content_begin,len(main_col)):
            content = self.sheet_data.row_values(i)
            if main_col[i] != "":
                id = x2jutils.getValueByType(main_col[i],main_type)
                table[id] = {}
                for k in range(0,len(self.types)):
                    if self.types[k] != '' and content[k] != '':
                        value = x2jutils.getValueByType(content[k],self.types[k],self.subTypes[k])
                        key = self.titles[k]
                        table[id][key] = value
        return table

    # 将单个sheet输出为一个字典, 根据标记位置, 将多行组成一个数组. 只支持2个level,即[{ 主键 [{从属内容},] },]
    def readExcelByGroup(self,findGroup,find):
        table = {}
        main_group_col = self.sheet_data.col_values(findGroup)
        main_col = self.sheet_data.col_values(find)
        main_group_type = self.types[findGroup]
        main_type = self.types[find]
        for i in range(self.content_begin,len(main_group_col)):
            content = self.sheet_data.row_values(i)
            if main_group_col[i] != "":
                group = x2jutils.getValueByType(main_group_col[i],main_group_type)
                if not group in table.keys():
                    table[group] = {}
                id = x2jutils.getValueByType(main_col[i],main_type)
                table[group][id] = {}
                for k in range(0,len(self.types)):
                    if self.types[k] != '' and content[k] != '':
                        value = x2jutils.getValueByType(content[k],self.types[k],self.subTypes[k])
                        key = self.titles[k]
                        table[group][id][key] = value   
        return table

    # **多语言文本表** 读取单个sheet, 以首列为字典key, 后续每一列作为不同dict的value, 并以每一列的id作为json名输出
    def readLocalizationExcel(self,char_data=None):
        keys = self.sheet_data.col_values(0)[self.content_begin:]
        col_num = self.sheet_data.ncols
        for i in range(1,col_num):
            if self.titles[i].startswith('#'):
                continue
            content = self.sheet_data.col_values(i)[self.content_begin:]
            if char_data:
                list_badChar = char_data.col_values(0)[1:]
                # list_badChar_unicode = char_data.col_values(1)[1:]
                list_goodChar = char_data.col_values(2)[1:]
                # list_goodChar_unicode = char_data.col_values(3)[1:]
                content = x2jutils.fixBadChar(content,list_badChar,list_goodChar)
            table = dict(zip(keys,content))
            x2jutils.writeJsonFile(os.path.join(self.output_path,self.titles[i]+'.json'),table)

if __name__ == "__main__":
    ax = x2jcore()
    # 获得所有json目录路径
    all_json_path = x2jutils.getAllFolders(ax.save_path) 

    #获得所有excel文件路径
    all_xlsx_path = []
    xlsx_folder_list = x2jutils.getAllFolders(ax.excel_path)
    for a in xlsx_folder_list:
        all_xlsx_path += [os.path.join(a,b) for b in x2jutils.xlsxFileList(a)]

    #清理temp目录
    x2jutils.clearTempFiles(ax.output_path)

    #导表
    for i in range(0,len(all_xlsx_path)):
        print(i+1,'---',all_xlsx_path[i])
    print('请输入要导出的表的序号:')
    selectFile=int(input())
    if selectFile>len(all_xlsx_path) or selectFile<1:
        print('输入错误 请重新执行脚本')
    else:
        ax.start(all_xlsx_path[selectFile-1])