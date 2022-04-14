import xlrd as xlrd
import os
from xlutils.copy import copy
from utils.langconv import Converter


# 繁体转简体
def TraditionalToSimplified(content):
    line = Converter("zh-hans").convert(content)
    return line


# 简体转繁体
def SimplifiedToTraditional(content):
    line = Converter("zh-hant").convert(content)
    return line


# Excel读取，写入，保存等操作类

# 获取存储行政区划数据的excel
def createExcel():
    # excel存储路径，绝对或相对路径都可以
    global excelPath
    excelPath = 'excel/result.xls'
    # excel文件
    global book
    book = xlrd.open_workbook(excelPath)
    # 用xlutils提供的copy方法将xlrd的对象转化为xlwt的对象
    global excel
    excel = copy(book)
    # 表格名称
    global countryTable
    global provinceTable
    # 获取存储国家信息表格
    countryTable = excel.get_sheet(0)
    provinceTable = excel.get_sheet(1)
    # 列名
    col = ('行政区划', '简介', '外交', '军事')
    for i in range(0, 4):
        countryTable.write(0, i, col[i])
        provinceTable.write(0, i, col[i])

    # 写入excel的起始下标
    global countryStartIndex
    global provinceStartIndex
    countryStartIndex = readLastIndex(excelPath, '国家')
    provinceStartIndex = readLastIndex(excelPath, '省份')


# 往excel中写入数据，
# kind：写入类型,0.往国家表中写入;1.往省份表中写入
# row:行，
# col:列，
# content:内容
def writeToExcel(kind, row, col, content):
    if kind == 0:
        countryTable.write(countryStartIndex + row, col, TraditionalToSimplified(content))
    elif kind == 1:
        provinceTable.write(provinceStartIndex + row, col, TraditionalToSimplified(content))


# 保存excel内容
def saveExcel():
    excel.save(excelPath)


# 读取excel表格最后一行数据行下标
# excelPath：表格的路径
# tableName：表的名称
def readLastIndex(excelPath, tableName):
    if not os.path.exists(excelPath):
        return 0
    # 首先打开excel表，formatting_info=True 代表保留excel原来的格式
    xls = xlrd.open_workbook(excelPath, formatting_info=False)
    # 通过sheet的名称获得sheet对象
    sheetFile = xls.sheet_by_name(tableName)
    return sheetFile.nrows - 1


# 读取excel表格指定行，列内的内容
# excelPath：表格的路径
# tableName：表的名称
# row:行
# rol:列
# return：返回指定行，列内的内容
def readRowColContent(excelPath, tableName, row, col):
    if not os.path.exists(excelPath):
        return None
    # 首先打开excel表，formatting_info=True 代表保留excel原来的格式
    xls = xlrd.open_workbook(excelPath, formatting_info=False)
    # 通过sheet的名称获得sheet对象
    sheetFile = xls.sheet_by_name(tableName)
    return sheetFile.row_values(row)[col]
