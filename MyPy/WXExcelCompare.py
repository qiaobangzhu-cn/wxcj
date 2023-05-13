# -*- coding:utf8 -*-

#合并两个采集结果，取出重合总数

import xlwings as xw
from config import quchong

xls_path1 = r'/Users/ruijieqiao/Desktop/生财有术采集/生财有术wx.xls'
xls_path2 = r'/Users/ruijieqiao/Desktop/淘金之路采集/淘金之路wx.xls'


def ReadExcel(xls_path):
    WXList = []
    print(xls_path)
    # 打开Excel程序，默认设置：程序可见，只打开不新建工作薄，屏幕更新关闭
    app = xw.App(visible=True, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    # 文件位置：xlsx_path，打开文档，然后保存，关闭，结束程序
    wb = app.books.open(xls_path)

    sheet1 = wb.sheets['sheet1']
    # rangeValue=sheet1.range('A2').value
    # print(rangeValue)

    rng = sheet1.range('a1').expand()
    nrows = rng.rows.count
    # 总行数
    print("总行数:" + str(nrows))

    # 循环遍历一列
    # for nrow in range(nrows+1):
    #    if nrow == 0:
    #        continue
    #    rangeValueTem = sheet1.range(f'A{nrow}').value
    #    print(rangeValueTem)

    # 一列数据
    AS = sheet1.range(f'a1:a{nrows}').value

    # 循环遍历这一列数据
    for a in AS:
        WXList.append(a)

    print(WXList)
    print(str(len(WXList)))

    wb.save()
    wb.close()
    app.quit()

    return WXList


WXList1 = ReadExcel(xls_path1)
SCcount = len(WXList1)
print("生财有术总数："+ str(SCcount))
WXList2 = ReadExcel(xls_path2)
TJcount = len(WXList2)
print("淘金之路总数："+ str(TJcount))

for wx in WXList2:
    WXList1.append(wx)

Acount1 = len(WXList1)
print("合并后的总数：" + str(Acount1))

wxlist = quchong(WXList1)
Acount2 = len(wxlist)
print("去重后的总数：" + str(Acount2))
print("差值："+str(Acount1-Acount2))