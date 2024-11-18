# 简介

常用于游戏研发的xlsx/xls转json数据的工具。

- [X]  **同时支持xlsx和xls格式**
- [X]  **一键移动到指定目录**
- [X]  **多语言文本特殊字符导出自动替换：** 需要自行补充替换文本
- [X]  **可导出多种json结构：** 可自行在x2jcore增加支持的种类
- [ ]  **多语言自动翻译：** 待开发

# 示例模板

请参考`配置表test/fortest.xlsx`和`配置表test/template.xlsx`

# 目录结构

`配置表/`  放置excel文件（可以有多个，配置表为前缀，例: 配置表_a 配置表临时）

`json/`  工具识别的json存放路径（可以有多个，json为前缀，例: json_A jsons4langs）

`output_temp/`  临时存放目录，工具每次导表会自动清除

`x2j工具`

# 运行

单文件版本：自行选择对应平台下载解压运行即可
[下载](https://github.com/dethanzhang/x2j/releases/tag/release)



python环境代码运行：请先安装依赖库xlrd，版本必须为1.2.0！

```bash
python3 -m pip install xlrd==1.2.0
```
