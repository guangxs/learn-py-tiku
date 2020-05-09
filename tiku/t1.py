import json

import requests
import xlrd
import xlwt
from xlutils.copy import copy
import re
from urllib import request
from urllib import error

import time

cou = '0'

book_name_xls = '超星题库'

value_title = [["题号", "课程ID", "题型", "题目", "答案"],]

urlInit = 'https://mooc1.chaoxing.com/course/{{courseId}}.html'
urlK = 'https://mooc1.chaoxing.com/nodedetailcontroller/visitnodedetail?courseId={{courseId}}&knowledgeId={{knowledgeId}}'
workUrl = 'https://mooc1.chaoxing.com/api/selectWorkQuestion?workId={{workId}}&ut=null&classId=0&courseId={{courseId}}&utenc=null'

headers = {
    'User-Agent': r'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) '
                  r'Chrome/45.0.2454.85 Safari/537.36 115Browser/6.0.3'
}
answerUrl = 'http://47.112.247.80/wkapi.php?q='

def __returnWorkUrl( courseId, workId):
    url = workUrl.replace('{{courseId}}', courseId).replace(
        '{{workId}}', workId)
    return url

def __getRequest( url):
    req = request.Request(url, headers=headers)
    try:
        page = request.urlopen(req).read()
        page = page.decode('utf-8')
        return page
    except error.URLError as e:
        print('courseId可能不存在哦！', e.reason)


def __getFristData(courseId):
    # 组装初始URL，获取第一个包含knowledge
    url = urlInit.replace('{{courseId}}', courseId)

    htmls = __getRequest(url)

    # <a class="wh nodeItem"  href="?courseId=200080607&knowledgeId=102433017" data="102433017">
    #re_rule = 'courseId='+courseId+'&knowledgeId=(.*)">'

    # <div id="" class="ml20 mb5  bgf3  mr10" data="102432997">
    # <div id="courseChapterSelected" class="ml20 bbe pl0 bg1e mr10" data="102433017" rel="">

    re_rule = 'courseId='+courseId+'&knowledgeId=(.*)">'
    url_frist = re.findall(re_rule, htmls)

    if len(url_frist) > 0:
        return url_frist[0]
    else:
        print(courseId, 'courseId错误！')

def __returnTitle(courseId, knowledgeId):
    url = urlK.replace('{{courseId}}', courseId).replace(
        '{{knowledgeId}}', knowledgeId)
    htmls = __getRequest(url)

    re_rule = '&quot;:&quot;work-(.*?)&quot;'
    wordId = re.findall(re_rule, htmls)
    wordId = list(set(wordId))  # 先转集合，再转队列  去重复

    title = []
    for x in wordId:
        wordUrl = __returnWorkUrl(courseId, x)
        html_work = __getRequest(wordUrl)
        title_rule = '<div class="Zy_TItle clearfix">\s*<i class="fl">.*</i>\s*<div class=".*">(.*?)</div>'
        title = title + re.findall(title_rule, html_work)

    # <div id="" class="ml20 mb5  bgf3  mr10" data="102432997">
    # <div id="courseChapterSelected" class="ml20 bbe pl0 bg1e mr10" data="102433017" rel="">

    #re_rule = '<a class=".*"  href="\?courseId=' + courseId+'&knowledgeId=.*" data="(.*)">'
    # re_rule = '<div id="(courseChapterSelected)?" class="[\s\S]*?" data="(\d*)">?'
    re_rule = '<div id="c?o?u?r?s?e?C?h?a?p?t?e?r?S?e?l?e?c?t?e?d?" class="[\s\S]*?" data="(\d*)">?'
    datas = re.findall(re_rule, htmls)

    return(title, datas)

def getTextByCourseId(courseId):
    global  cou
    #try:
    titles = []
    data_now = __getFristData(courseId)  # 第一个data需要再单独的一个链接里获取
    j = 1
    while data_now:
        listR = __returnTitle(courseId, data_now)

        title = listR[0]
        data = listR[1]

        for i, x in enumerate(data):
            if data_now == x:
                if len(data) > (i+1):
                    data_now = data[i+1]
                else:
                    data_now = None
                    print('获取题目结束.')
                break

        # 打印题目  去除题目中的<p></p>获取其他标签，只有部分题目有，可能是尔雅自己整理时候加入的。
        p = re.compile(r'[【](.*?)[】]', re.S)
        for t in title:
            p_rule = '<.*?>'
            t = re.sub(p_rule, '', t)
            p_rule = '&.*?;'
            t = re.sub(p_rule, '', t)

            r = requests.post(answerUrl + t)
            rd = json.loads(r.text)
            k = [j, courseId, re.findall(p, t), t, rd['answer']]
            titles.append(k)
            if cou != courseId:
                write_excel_xls(book_name_xls + courseId + '.xls', courseId, value_title)
                write_excel_xls_append(book_name_xls + courseId + '.xls', k)
                cou = courseId
            else:
                write_excel_xls_append(book_name_xls + courseId + '.xls', k)
            print(k)
            j += 1
    return titles
#except:
#return []


def write_excel_xls(path, sheet_name, value):
    index = len(value)  # 获取需要写入数据的行数
    workbook = xlwt.Workbook()  # 新建一个工作簿
    sheet = workbook.add_sheet(sheet_name)  # 在工作簿中新建一个表格
    for i in range(0, index):
        for j in range(0, len(value[i])):
            sheet.write(i, j, value[i][j])  # 像表格中写入数据（对应的行和列）
    workbook.save(path)  # 保存工作簿
    print("xls格式表格写入数据成功！")


def write_excel_xls_append(path, value):
    index = len(value)  # 获取需要写入数据的行数
    workbook = xlrd.open_workbook(path)  # 打开工作簿
    sheets = workbook.sheet_names()  # 获取工作簿中的所有表格
    worksheet = workbook.sheet_by_name(sheets[0])  # 获取工作簿中所有表格中的的第一个表格
    rows_old = worksheet.nrows  # 获取表格中已存在的数据的行数
    new_workbook = copy(workbook)  # 将xlrd对象拷贝转化为xlwt对象
    new_worksheet = new_workbook.get_sheet(0)  # 获取转化后工作簿中的第一个表格

    for j in range(0, len(value)):
        new_worksheet.write(rows_old, j, value[j])  # 追加写入数据，注意是从i+rows_old行开始写入
    new_workbook.save(path)  # 保存工作簿
    print("xls格式表格【追加】写入数据成功！")


def read_excel_xls(path):
    workbook = xlrd.open_workbook(path)  # 打开工作簿
    sheets = workbook.sheet_names()  # 获取工作簿中的所有表格
    worksheet = workbook.sheet_by_name(sheets[0])  # 获取工作簿中所有表格中的的第一个表格
    for i in range(0, worksheet.nrows):
        for j in range(0, worksheet.ncols):
            print(worksheet.cell_value(i, j), "\t", end="")  # 逐行逐列读取数据


#input('请输入courseId:')  # '200837021'  200080607 = 189题
if __name__ == '__main__':

    i = 208422029
    while(i):
        #write_excel_xls(book_name_xls, '1', value_title)

        try:
            courseId = str(i)
            getTextByCourseId(courseId)
        except:
            print("数据异常")
            time.sleep(1) # 暂停 1 秒
        i += 1
        if i > 400000000:
            exit()
