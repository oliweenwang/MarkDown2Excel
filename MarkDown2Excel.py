#!/usr/bin/python3
# coding:utf-8
# python3.6.4

import xlsxwriter
import string

# 定义一个列表保存Excel列编号
columnList = list(string.ascii_uppercase)

fileName = '健客网'
baseDir = 'C:/Users/Ray/Desktop/'
markDownName = baseDir + fileName +'.md'
excelName = baseDir + fileName +'.xlsx'

# 打开gbk编码的MarkDown文件
markDownFile = open(markDownName, 'r', encoding='utf-8')
# 创建一个Excel表格
excelFile = xlsxwriter.Workbook(excelName)
# 再表格中创建一张表
worksheet = excelFile.add_worksheet()
# 定义一个加粗和增加自动换行格式
bold = excelFile.add_format({'bold': True})
bold.set_text_wrap()
# 定义一个自动换行格式
autoWrap = excelFile.add_format()
autoWrap.set_text_wrap()

# 设置工作表列宽(单位不是像素)
worksheet.set_column('A:A', 5)
worksheet.set_column('B:B', 11)
worksheet.set_column('C:C', 80)
worksheet.set_column('D:D', 15)
worksheet.set_column('E:E', 20)

# MarkDown文件的行编号
lineNum = 1
# MarkDown文件的列编号
columnNum = 0

# 处理MarkDown文件
for line in markDownFile.readlines():
    line = line.replace(' ', '').replace('\n', '').replace('<br>','\n').split('|')
    # print(type(line), line)
    # 处理标题行
    if line[0] == '序号':
        for i in line:
            worksheet.write(columnList[columnNum] + str(lineNum), i, bold)
            columnNum += 1
        lineNum += 1
    # 处理内容行
    elif line[0].isalnum():
        # print(line)
        for i in line:
            if columnNum == 0:
                i = int(i)
                worksheet.write_number(columnList[columnNum] + str(lineNum), i, autoWrap)
            elif columnNum == 3:
                i = float(i)
                worksheet.write_number(columnList[columnNum] + str(lineNum), i, autoWrap)
            else:
                worksheet.write(columnList[columnNum] + str(lineNum), i, autoWrap)
            columnNum += 1
        lineNum += 1
    columnNum = 0

markDownFile.close()
excelFile.close()
