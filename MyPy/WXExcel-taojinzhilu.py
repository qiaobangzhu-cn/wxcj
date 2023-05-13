# -*- coding:utf8 -*-

#将多个Excel的结果去重合并

from config import ReadExcel
from config import WriteExcel

xlsx_path1 = r'/Users/ruijieqiao/Desktop/淘金之路采集/淘金之路采集.xlsx'
xlsx_path2 = r'/Users/ruijieqiao/Desktop/淘金之路采集/淘金之路俱乐部采集.xlsx'
xlsx_path3 = r'/Users/ruijieqiao/Desktop/淘金之路采集/淘金之路资源对接采集.xlsx'
#需要新建wx.xls文件，否则报错
xls_path1 = r'/Users/ruijieqiao/Desktop/淘金之路采集/淘金之路wx.xls'
xls_path2 = r'/Users/ruijieqiao/Desktop/淘金之路采集/微信xls/淘金之路wx'

# 将WXList按照等分切割，即一个xls文件中需要包含的记录条数
ROW = 999

ReadExcel(xlsx_path1)
ReadExcel(xlsx_path2)
ReadExcel(xlsx_path3)
WriteExcel(xls_path1)
#SplitExcel()
