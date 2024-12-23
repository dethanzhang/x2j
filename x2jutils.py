# !/usr/bin/python3
# -*- coding: utf-8 -*-
import os
import json


def checkChdir():  # 打包为mac单文件时, 需要切换到app的上一级目录
    now = os.getcwd()
    if ".app" in now:
        os.chdir(os.path.split(now.split(".app")[0])[0])


def fileExtension(path):
    return os.path.splitext(path)[1]


def getAllFolders(prefix):
    return [x for x in os.listdir() if x.startswith(prefix) and os.path.isdir(x)]


def clearTempFiles(filePath):
    if os.path.exists(filePath):
        # 如果存在的话清空里面所有json/txt文件
        list_files = os.listdir(filePath)
        for i in range(0, len(list_files)):
            file = list_files[i]
            subfilePath = os.path.join(filePath, file)
            if os.path.isfile(subfilePath) and (
                fileExtension(file) == ".json" or fileExtension(file) == ".txt"
            ):
                os.remove(subfilePath)
    else:
        os.makedirs(filePath)


def xlsxFileList(filePath):
    listDir = sorted(os.listdir(filePath), key=len)
    filenames = []
    for x in listDir:
        if (not x.startswith("~$")) and (
            fileExtension(x) == ".xlsx" or fileExtension(x) == ".xls"
        ):
            filenames.append(x)
    return filenames


def autoMove(dir1, dir2):
    list_files = os.listdir(dir1)
    for x in list_files:
        try:
            if os.path.exists(os.path.join(dir2, x)):
                os.remove(os.path.join(dir2, x))
            os.rename(os.path.join(dir1, x), os.path.join(dir2, x))
        except Exception as e:
            return str(e)
    return None


def writeJsonFile(filePath, data):
    data_str = json.dumps(data, sort_keys=False, indent=4, ensure_ascii=False)
    if os.path.exists(filePath):
        os.remove(filePath)
    with open(filePath, "w", encoding="utf8") as f:
        f.write(data_str)
    if os.name == "posix":  # 非windows系统执行, 将LF转换为CRLF
        with open(filePath, "rb") as f:
            content = f.read()
        content = content.replace(b"\n", b"\r\n")
        with open(filePath, "wb") as f:
            f.write(content)


def fixBadChar(content, list_badChar, list_goodChar):
    for i, x in enumerate(list_badChar):
        for idx, c in enumerate(content):
            if x in str(c):
                content[idx] = c.replace(x, list_goodChar[i])
                print("检测到字符错误, 已替换为", list_goodChar[i])
    return content


def autoValue(value):
    try:
        value = float(value)
    except ValueError:
        return value
    value = f"{value:.8f}".rstrip("0").rstrip(".")
    try:  # 是否整数
        return int(value)
    except ValueError:
        pass
    return float(value)


def getValueByType(value, type1, subType=None, debug=False):
    try:
        # 处理填null的空值
        if value == "null":
            if (
                type1 == "matrix"
                or type1 == "array"
                or type1 == "array-str"
                or type1 == "json"
                or type1 == "json-str"
            ):
                return []
            if type1 == "int" or type1 == "float":
                return 0
            return ""

        # 处理int/float
        if type1 == "int":
            if value == "":
                return 0
            return int(value)
        if type1 == "float":
            return float(
                autoValue(value)
            )  # 直接导出会导致导出的值为内存存储值, 与实际显示不一致

        # 处理bool
        if type1 == "bool":
            if (
                value == "true"
                or autoValue(value) == 1
                or value == "True"
                or value == "TRUE"
            ):
                return True
            return False

        # 处理array-str，即所有数组内元素都为字符串
        if type1 == "array-str":
            if "\n" in str(value):
                value = value.replace("\n", ",")
            if "," not in str(value):
                return [str(autoValue(value))]
            listsValue = [str(autoValue(i)) for i in value.split(",")]
            return listsValue

        # 处理array，数组内各元素的类型自动识别
        if type1 == "array":
            if "\n" in str(value):
                value = value.replace("\n", ",")
            if value == "":
                return []
            if "," not in str(value):
                value = autoValue(value)
                return [value]
            listsValue = [autoValue(i) for i in value.split(",")]
            return listsValue

        # 处理matrix二维数组，各元素类型自动识别
        if type1 == "matrix":
            if "\n" in str(value):
                value = value.replace("\n", "|")
            if "|" not in str(value):
                value = getValueByType(value, "array")
                return [value]
            matrixList = [x for x in value.split("|")]
            for i in range(len(matrixList)):
                matrixList[i] = getValueByType(matrixList[i], "array")
            return matrixList

        # 处理matrix-str二维数组，各元素类型均转为字符串
        if type1 == "matrix-str":
            if "\n" in str(value):
                value = value.replace("\n", "|")
            if "|" not in str(value):
                value = getValueByType(value, "array-str")
                return [value]
            matrixList = [x for x in value.split("|")]
            for i in range(len(matrixList)):
                matrixList[i] = getValueByType(matrixList[i], "array-str")
            return matrixList

        # 处理json-str，即所有json内元素都为字符串
        # 填写格式为：subType行填写k1,k2,k3  内容行填写v01,v02,v03;v11,v12,v13;...
        # 输出格式为： [{k1:v01,k2:v02,...},{k1:v11,k2:v12,...},...]
        if type1 == "json-str":
            if value == "":
                return []
            if subType:
                if "\n" in str(value):
                    value = value.replace("\n", ";")
                maps = subType.split(",")
                strs = value.split(";")
                ret = []
                if value == "":
                    return ret
                for i in range(0, len(strs)):
                    dic = {}
                    contents = strs[i].split(",")
                    for j in range(0, len(maps)):
                        dic[maps[j]] = getValueByType(contents[j], "str")
                    ret.append(dic)
                return ret
            else:
                return json.loads(value)

        # 处理json，各元素类型自动识别
        if type1 == "json":
            if value == "":
                return []
            if subType:
                if "\n" in str(value):
                    value = value.replace("\n", ";")
                maps = subType.split(",")
                strs = value.split(";")
                ret = []
                for i in range(0, len(strs)):
                    dic = {}
                    contents = strs[i].split(",")
                    for j in range(0, len(maps)):
                        dic[maps[j]] = autoValue(contents[j])
                    ret.append(dic)
                return ret
            else:
                return json.loads(value)

        # 处理str/string，字符串类型自动识别
        if type1 == "str" or type1 == "string":
            if type(value).__name__ == "str":
                return value
            value = autoValue(value)
            return str(value)

    except Exception as e:
        print("[ERROR]getValueByType >> %s" % e)
        if debug:
            return e
        return None
