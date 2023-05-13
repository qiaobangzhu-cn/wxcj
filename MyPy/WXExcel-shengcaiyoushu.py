# -*- coding:utf8 -*-

#将多个Excel的结果去重合并

from io import StringIO
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.pdfinterp import process_pdf
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
import re

from config import quchong
from config import ReadExcel
from config import WriteExcel
from config import gloVar

xlsx_path1 = r'/Users/ruijieqiao/Desktop/生财有术采集/微信ID采集.xlsx'
xlsx_path2 = r'/Users/ruijieqiao/Desktop/生财有术采集/资源对接采集1224.xlsx'
xlsx_path3 = r'/Users/ruijieqiao/Desktop/生财有术采集/生财有术最新采集.xlsx'
#需要新建wx.xls文件，否则报错
xls_path1 = r'/Users/ruijieqiao/Desktop/生财有术采集/生财有术wx.xls'
xls_path2 = r'/Users/ruijieqiao/Desktop/生财有术采集/微信xls/生财有术wx'
pdf_path1 = r'/Users/ruijieqiao/Desktop/生财有术采集/生财有术之升级朋友圈.pdf'


# 将WXList按照等分切割，即一个xls文件中需要包含的记录条数
ROW = 999


def ReadPDF():
    global pdf_path1
    print(pdf_path1)
    print("WXList")
    print(gloVar.WXList)
    print(len(gloVar.WXList))
    with open(pdf_path1, 'rb') as file:
        resource_manager = PDFResourceManager()
        return_str = StringIO()
        lap_params = LAParams()
        device = TextConverter(resource_manager, return_str, laparams=lap_params)
        process_pdf(resource_manager, device, file)
        device.close()
        content = return_str.getvalue()
        return_str.close()
        #        return re.sub('\s+', '', content)
        # 先过滤出
        rs = re.findall("[】|】：][a-zA-Z\d][a-zA-Z\d_-]{5,19}", content)
        for r in rs:
            gloVar.WXList.append(re.findall("[a-zA-Z\d][a-zA-Z\d_-]{5,19}", r)[0])

    print("增加后的WXList:")
    print(gloVar.WXList)
    print(str(len(gloVar.WXList)))
    print("增加后且去重的WXList：")
    gloVar.WXList = quchong(gloVar.WXList)
    print(gloVar.WXList)
    print(str(len(gloVar.WXList)))

#    for c in range(1,xls_count+1):
#        print(c)


ReadExcel(xlsx_path1)
ReadExcel(xlsx_path2)
ReadExcel(xlsx_path3)
ReadPDF()
WriteExcel(xls_path1)
#SplitExcel()
