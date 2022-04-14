# coding:utf-8
import operator

from upload import NetControl
from utils import ExcelParser
from utils.langconv import Converter


# 繁体转简体
def TraditionalToSimplified(content):
    line = Converter("zh-hans").convert(content)
    return line


# 简体转繁体
def SimplifiedToTraditional(content):
    line = Converter("zh-hant").convert(content)
    return line


# wiki数据分析

# 打印和存储日志
def record(content):
    print(content)
    logFile.write(TraditionalToSimplified(content) + '\n')


# 获取国家简介
def getIntroduce(kind, body, realCount):
    # 实际需要写入excel的内容
    content = ''
    # 获取国家简介
    for ps in body.contents:
        if ps.name is not None:
            if ps.name == 'p':
                for contents in ps.descendants:
                    if contents.name is not None:
                        if contents.name == 'style':
                            continue
                        if contents.string is not None:
                            wikiFile.write(TraditionalToSimplified(contents.string))
                            content += contents.string
            wikiFile.write("\n")
            content += "\n"
        # 检查国家是否获取完毕并终止查询
        next = ps.next_sibling
        if next is not None:
            if operator.contains(str(next), 'role="navigation"'):
                break

    wikiFile.write("\n\n\n\n\n")
    ExcelParser.writeToExcel(kind, realCount, 1, content)


# 获取外交描述
def getRelation(kind, body, realCount):
    # 实际需要写入excel的内容
    content = ''
    # 标识是否找到外交数据
    found = 0
    for ps in body.contents:
        if ps.name is not None and found == 0:
            if ps.name == 'h2':
                for head in ps.find_all('span', class_='mw-headline'):
                    if operator.contains(str(head), 'id="外交"'):
                        found = 1
        if found == 0:
            continue
        if found == 1:
            if ps.name is not None:
                for contents in ps.descendants:
                    if contents.string is not None:
                        wikiFile.write(TraditionalToSimplified(contents.string))
                        content += contents.string
                    else:
                        for child in contents.strings:
                            if child.string is not None:
                                wikiFile.write(TraditionalToSimplified(child.string))
                                content += child.string
                wikiFile.write("\n")
                content += "\n"

        # 检查外交是否获取完毕并终止查询
        next = ps.next_sibling
        if next is not None and next.name is not None:
            if next.name == 'h2':
                break

    wikiFile.write("\n\n\n\n\n")
    ExcelParser.writeToExcel(kind, realCount, 2, content)


# 获取军事描述
def getMilitary(kind, body, realCount):
    # 实际需要写入excel的内容
    content = ''
    # 标识是否找到外交数据
    found = 0
    for ps in body.contents:
        if ps.name is not None and found == 0:
            if ps.name == 'h2':
                for head in ps.find_all('span', class_='mw-headline'):
                    if operator.contains(str(head), 'id="军事"'):
                        found = 1
        if found == 0:
            continue
        if found == 1:
            if ps.name is not None:
                for contents in ps.descendants:
                    if contents.string is not None:
                        wikiFile.write(TraditionalToSimplified(contents.string))
                        content += contents.string
                    else:
                        for child in contents.strings:
                            if child.string is not None:
                                wikiFile.write(TraditionalToSimplified(child.string))
                                content += child.string
                wikiFile.write("\n")
                content += "\n"

        # 检查军事是否获取完毕并终止查询
        next = ps.next_sibling
        if next is not None and next.name is not None:
            if next.name == 'h2':
                break

    ExcelParser.writeToExcel(kind, realCount, 3, content)


# 核心代码：解析获取到的soup数据
# kind：爬取数据类型，0：爬取国家数据，1：爬取省份数据
# soup：抓取的标签主体
# targetName：行政区划名称
# wikiFile：存储抓取到维基百科信息的文件
# realCount:当前编写的行
def parseSoup(kind, soup, targetName, wikiFile, realCount):
    # 获取行政区划标题
    title = soup.find('h1').string
    wikiFile.write(title + "\n\n\n\n\n")
    ExcelParser.writeToExcel(kind, realCount, 0, title)

    record('开始解析：' + targetName)

    parent = soup.find('div', class_='mw-body-content mw-content-ltr')
    if parent is None:
        errorFile.write('---------' + targetName + '无可抓取信息---------\n')
        record('---------' + targetName + '无可抓取信息---------\n')
        return
    body = parent.find('div', class_='mw-parser-output')
    if body is None:
        errorFile.write('---------' + targetName + '无可抓取信息---------\n')
        record('---------' + targetName + '无可抓取信息---------\n')
        return

    # 获取国家简介
    getIntroduce(kind, body, realCount)

    # 获取外交描述
    getRelation(kind, body, realCount)

    # 获取军事描述
    getMilitary(kind, body, realCount)

    record('解析并存储 ' + targetName + ' 成功！\n')


# 读取行政区划信息
# jumpIndex：跳过前面jumpIndex行数据再进行抓取，避免网络中断或其他情况导致数据重新从头开始抓取
# endIndex：到endIndex行时终止抓取
# maxRow：excel里面有内容的总行数
def parseExcel(jumpIndex, endIndex, maxRow):
    if endIndex != -1:
        if jumpIndex > endIndex:
            jumpIndex = endIndex
        if endIndex > maxRow:
            endIndex = maxRow
    else:
        if jumpIndex > maxRow:
            jumpIndex = maxRow
    lastName = ''
    row = 1
    countryCount = 0
    provinceCount = 0
    while row <= maxRow:
        if endIndex != -1:
            if row > endIndex:
                break
        # 跳过前面jumpIndex条数据
        if row < jumpIndex:
            record('跳过第' + str(row) + '条的抓取\n')
            row += 1
            continue
        # 找到需要爬取的行政区划名称
        countryName = ExcelParser.readRowColContent(excelPath, tableName, row, 0)
        if countryName is None:
            record('未找到文件：' + excelPath + '\n')
            row += 1
            continue
        if countryName == '':
            record('读取' + tableName + '第' + row + '行，第' + str(0) + '列数据为空\n')
            row += 1
            continue

        # 爬取国家数据
        if lastName != countryName:
            record(
                '抓取第 ' + str(row) + ' 行数据，第 ' + str(countryCount + 1) + ' 条国家  ********  ' + countryName + '  ********')
            fileUrl = wikiUrl + countryName
            soup = NetControl.sendRequest(fileUrl)
            if soup is None:
                row += 1
                lastName = countryName
                errorFile.write('网络请求异常，' + countryName + ' 请求失败\n')
                record('网络请求异常，' + countryName + ' 请求失败\n')
                continue
            countryCount += 1
            parseSoup(0, soup, countryName, wikiFile, countryCount)

        # 爬取省份数据
        provinceName = ExcelParser.readRowColContent(excelPath, tableName, row - 1, 1)
        record('抓取第 ' + str(row) + ' 行数据，省份 ： ' + provinceName)
        fileUrl = wikiUrl + provinceName
        soup = NetControl.sendRequest(fileUrl)
        if soup is None:
            row += 1
            errorFile.write('网络请求异常，' + provinceName + ' 请求失败\n')
            record('网络请求异常，' + provinceName + ' 请求失败\n')
            continue
        provinceCount += 1
        parseSoup(1, soup, provinceName, wikiFile, provinceCount)

        row += 1
        lastName = countryName


# 全部需要爬取的国家，省份的行政区划，绝对或相对路径都可以
excelPath = 'excel/countrys.xls'

# 表格的名称
tableName = '行政区划'

# 维基百科抓取的数据本地文件记录
wikiFile = open('record/wiki.txt', mode='w', encoding='utf-8')

# 抓取过程的全部记录信息
logFile = open('record/log.txt', mode='w', encoding='utf-8')

# 抓取过程的全部错误记录信息
errorFile = open('record/error.txt', mode='w', encoding='utf-8')

# 维基百科网络地址
wikiUrl = 'https://ww.wiki.fallingwaterdesignbuild.com/wiki/'


# 执行程序
# jumpIndex：跳过前面jumpIndex行数据再进行抓取，避免程序异常中断重新从头开始抓取，下标从0开始
# 不需要时，传入0即可，需要注意的是，传入的是excel的实际行数，不是第多少个行政区划
# endIndex：到endIndex行时终止抓取，不需要时候，传入-1即可，
# 需要注意的是，传入的是excel的实际行数，不是第多少个行政区划
def run(jumpIndex, endIndex):
    if jumpIndex < 1:
        jumpIndex = 1
    if endIndex < 0:
        endIndex = -1

    record('\n*******      打开文件：' + excelPath + '，并开始解析表：' + tableName + '      *******\n')

    # 创建excel类型文件及其相应表格，列名
    ExcelParser.createExcel()

    if endIndex == -1:
        record('解析从 ' + str(jumpIndex) + ' 行开始到解析完整个文件\n')
    else:
        record('解析从 ' + str(jumpIndex) + ' 行开始,到 ' + str(endIndex) + ' 行终止\n')
    # 获取全部需要爬取的国家行数
    maxRow = ExcelParser.readLastIndex(excelPath, tableName)

    # 爬取国家信息
    parseExcel(jumpIndex, endIndex, maxRow)

    # 保存excel
    ExcelParser.saveExcel()

    record('*******      爬取结束，关闭文件：' + excelPath + '      *******')


if __name__ == '__main__':
    run(1,-1)
