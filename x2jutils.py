# !/usr/bin/python3
# -*- coding: utf-8 -*-

import os
import json

CONST_EXCEL_PATH = "配置表utils"
CONST_OUTPUT_PATH = 'output_temp'
CONST_SAVE_PATH = 'json'
CONST_ERROR_MSG = []

def checkChdir():
    now = os.getcwd()
    if '.app' in now:
        os.chdir(os.path.split(now.split('.app')[0])[0])

def fileExtension(path): 
    return os.path.splitext(path)[1]

def clearTempFiles(filePath):
    if os.path.exists(filePath):
        #如果存在的话清空里面所有json/txt文件
        list_files = os.listdir(filePath)
        for i in range(0,len(list_files)):
            file = list_files[i]
            subfilePath = os.path.join(filePath,file)
            if os.path.isfile(subfilePath) and (fileExtension(file) == '.json' or fileExtension(file) == '.txt'):
                os.remove(subfilePath)
    else:
        os.makedirs(filePath)

def getAllFolders(prefix):
    return [x for x in os.listdir() if x.startswith(prefix) and os.path.isdir(x)]

def pathFileList(filePath):
    listDir = sorted(os.listdir(filePath))
    filenames = []
    for x in listDir:
        if (not x.startswith('~$')) and (fileExtension(x) == '.xlsx' or fileExtension(x) == '.xls'):
            filenames.append(x)
    return filenames

def autoMove(dir1,dir2):
    list_files = os.listdir(dir1)
    for x in list_files:
        try:
            if os.path.exists(os.path.join(dir2,x)):
                os.remove(os.path.join(dir2,x))
            os.rename(os.path.join(dir1,x),os.path.join(dir2,x))
        except Exception as e:
            CONST_ERROR_MSG.append(str(e))
            return e


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

def fixBadChar(content,list_badChar,list_goodChar):
    for i,x in enumerate(list_badChar):
        for idx,c in enumerate(content):
            if x in str(c):
                content[idx] = c.replace(x,list_goodChar[i])
                print("检测到字符错误, 已替换为",list_goodChar[i])
    return content

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

def getValueByType(value,type1,subType=None,debug=False):
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
                value=value.replace('\n','|')
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

    except Exception as e:
        print('[ERROR]getValueByType >> %s'%e)
        # CONST_ERROR_MSG.append("第【%d】行 第【%d】列单元格填写错误"%(row+1,column+1))
        if debug:
            return e
        return None