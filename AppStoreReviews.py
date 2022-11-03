#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Oct 31 19:31:48 2022

@author: hq
"""

import urllib.request
import xlsxwriter
import json
import os

def getResponseJson(url):
    response = urllib.request.urlopen(url)
    responseJson = json.loads(response.read().decode())

    return responseJson

def main():
    appID = input('请输入AppID：')
    appName = input('请输入App名称：')
    totalNeed = input('共需要多少页评论（一页50条）：')

    if not os.path.exists(appID):
        os.system("mkdir " + appID)

    outputExcel = xlsxwriter.Workbook('./' + appID + '/' + appName + '.xlsx')
    outputExcelSheet = outputExcel.add_worksheet()

    format = outputExcel.add_format()
    format.set_border(1)

    formatTitle = outputExcel.add_format()
    formatTitle.set_border(1)
    formatTitle.set_align('left')
    formatTitle.set_bg_color('#ababab')
    formatTitle.set_bold()

    title = ['作者', '标题', '评论内容', '版本', '时间']


    outputExcelSheet.set_column(0, 0, 20)
    outputExcelSheet.set_column(1, 1, 20)
    outputExcelSheet.set_column(2, 2, 150)
    outputExcelSheet.set_column(3, 3, 20)
    outputExcelSheet.set_column(4, 4, 40)

    outputExcelSheet.write_row('A1', title, formatTitle)

    count = 0
    total = int(totalNeed)
    rowCount = 0

    for i in range(total):
        url = 'https://itunes.apple.com/rss/customerreviews/page=' + str(i+1)\
            + '/id=' + str(appID) + '/sortby=mostrecent/json?l=en&&cc=chn'

        jsonGet = getResponseJson(url)

        fileName = appID + '/' + str(i+1) + '.json'

        feed = jsonGet['feed']
        entry = feed['entry']

        for j in range(len(entry)):
            item = entry[j]
            startRow = rowCount + 1

            outputExcelSheet.write(
                startRow, 0, item['author']['name']['label'], format)
            outputExcelSheet.write(
                startRow, 1, item['title']['label'], format)
            outputExcelSheet.write(
                startRow, 2, item['content']['label'], format)
            outputExcelSheet.write(
                startRow, 3, item['im:version']['label'], format)
            outputExcelSheet.write(
                startRow, 4, item['updated']['label'], format)

            rowCount = rowCount + 1

        with open(fileName, 'w') as file:
            file.write(json.dumps(jsonGet, sort_keys=True,
                                  indent=4, ensure_ascii=False))

        count = count + 1
        print(str(count) + '/' + str(total))

    outputExcel.close()

main()
