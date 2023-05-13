# -*- coding:utf8 -*-

import xlwings as xw
import math
from io import StringIO
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.pdfinterp import process_pdf
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
import re


class gloVar():
    # 将Excel里面数据组装
    WXList = []

def quchong(arr):
    return list(set(arr))

#需要过滤掉的微信号
FilterWX = ['HK30-10', '30000X6', 'Building', 'Precios', 'webhook', 'hahaha', '201----', 'shopee', 'Facebook',
            'youtuber', 'window', 'soogif', 'caozhao', 'Helium', 'banner', 'L2B5200', 'FigmaChina', 'silence',
            'fackbook', 'kangkang', 'iconfont-', '8280000', 'Machinery', 'Personalized', 'wetool', 'listing',
            'TOASTMASTERS', 'chrome', 'alibaba', 'ProcessOn', 'Grammarly', '30000X3', 'douyin', 'windows', 'Releases',
            'maomao', 'summer', 'chatgpt', 'Envato', 'Bitcoin', 'google', 'livres', 'wechat', 'expression', 'Typora',
            'coulter', 'youtube', 'Poisoner', 'amazon', 'manchuan', 'Virtual', 'LookAE', 'Binance', 'xiaolai', 'selina',
            'thanks', '30000X0', 'leader', 'android', '666666', 'manychat', 'MoChat', 'vivian', 'filestorage',
            '90000X30', 'Dropshipping', 'discuzq', 'switch', 'Cryptocurrency', 'WeChatDownload', '180X30', '1200book',
            'slogan', 'compra', 'iPhone', 'airpods', 'Seeseed-', 'master', 'coupang', 'expires', 'website', 'Online',
            'cooking', 'Personal', 'Amazon', 'Combinator', 'Attention', '1075375006', 'featuers', 'review', 'whatspp',
            'Aimmon', '30000', '180000X30', 'python', 'whatsapp', 'Cloudflare', 'Traveling', '212400', 'shorts',
            'tiktok', 'Unlimited', 'shopify', '20201210', 'github', 'dreaming', 'Instagram', 'microsoft', 'Create',
            'vscode', 'robots', '66666666666', 'xxxxxx', 'mdnice', 'openao', 'elettronica', 'ArcTime','20220217',
            'SSSSSSS','biubiu','change']
TmpFilterWX = []

# 全部转小写
for f in FilterWX:
    TmpFilterWX.append(f.lower())

FilterWX = TmpFilterWX
#需要过滤掉的微信号的正则表达式
Pattern = ["\d{1,9}-\d{1,9}", "\d{1,9}[w|W]-\d{1,9}[w|W]", "\d{1,9}-\d{1,9}[a-z]{1,2}", "img_\d{1,5}"]




'''
print(filterWX)

w ="11A"
print(w.lower())

ff = re.match("\d{1,9}-\d{1,9}", "1112-111-111")
print(ff)
if ff != None:
    print(ff.group())

ff = re.match("\d{1,9}[w|W]-\d{1,9}[w|W]", "1112w-111w-111")
print(ff)
if ff != None:
    print(ff.group())

ff = re.match("\d{1,9}-\d{1,9}[a-z]{1,2}", "111-111cm")
print(ff)
if ff != None:
    print(ff.group())

ff = re.match("img_\d{1,5}", "img_1111")
print(ff)
if ff != None:
    print(ff.group())
'''

# 微信ID采集.xlsx处理
def ReadExcel(xlsx_path):
    print(xlsx_path)
    print("WXList:")
    print(gloVar.WXList)
    print(len(gloVar.WXList))

    # 打开Excel程序，默认设置：程序可见，只打开不新建工作薄，屏幕更新关闭
    app = xw.App(visible=True, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    # 文件位置：xlsx_path，打开文档，然后保存，关闭，结束程序
    wb = app.books.open(xlsx_path)

    sheet1 = wb.sheets['sheet1']
    # rangeValue=sheet1.range('A2').value
    # print(rangeValue)

    rng = sheet1.range('a1').expand()
    nrows = rng.rows.count
    # 总行数
    print("Excel总行数nrows:" + str(nrows))
    print(str(nrows))

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
        # 取每一列数据
        a = a.split("\r\n")
        for aTem in a:
            isWX = True
            #先转成小写
            aTemLower = aTem.lower()

            #遍历过滤器里面的内容
            for f in FilterWX:
                if f == aTemLower:
                    isWX = False
                    break

            #再判断正则匹配
            for p in Pattern:
                r = re.match(p,aTemLower)
                if r != None:
                    isWX = False

            if isWX:
                gloVar.WXList.append(aTem)
            else:
                print("不符合的微信："+ aTem)

    print(xlsx_path)
    print("Excel取出来的值AS：")
    print(AS)
    print(len(AS))
    print("经过过滤筛选后的值WXList:")
    print(gloVar.WXList)
    print(str(len(gloVar.WXList)))
    print("经过筛选后且去重的值WXList：")
    gloVar.WXList = quchong(gloVar.WXList)
    print(gloVar.WXList)
    print(str(len(gloVar.WXList)))

    wb.save()
    wb.close()
    app.quit()


# 写入新的xls文件
def WriteExcel(xls_path):
    # 打开Excel程序，默认设置：程序可见，只打开不新建工作薄，屏幕更新关闭
    app = xw.App(visible=True, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    # 文件位置：xlsx_path，打开文档，然后保存，关闭，结束程序
    wb = app.books.open(xls_path)
    sheet1 = wb.sheets['sheet1']
    # 先清空之前记录
    sheet1.clear()
    sheet1.range('A1').value = "微信号/手机号/QQ号"
    sheet1.range('A2').options(transpose=True).value = gloVar.WXList

    wb.save()
    wb.close()
    app.quit()


# 将WXList内容按照等分切割到不同xls文件中
def SplitExcel():
    global ROW
    global WXList
    global xls_path2
    WXList_count = len(WXList)
    # 向上取整
    # xls_count为总共生成的文件数
    xls_count = math.ceil(WXList_count / ROW)
    print(xls_count)
    c = 0;
    # 按照等分数遍历，把等分数据写入不同文件中
    while c != xls_count:
        c = c + 1
        print(c)
        xls_file_name = None
        app = xw.App(visible=True, add_book=False)
        wb = app.books.add()
        sheet1 = wb.sheets['sheet1']
        sheet1.name = "数据详情"
        sheet1.range('A1').value = ["手机/QQ/微信", "备注", "验证语"]

        if c == 1:
            xls_file_name = xls_path2 + "1-" + str(ROW) + ".xls"
            print(WXList[0:ROW])
            print(xls_file_name)
            sheet1.range('A2').options(transpose=True).value = WXList[0:ROW]
        elif c == xls_count:
            xls_file_name = xls_path2 + str((c - 1) * ROW) + "-" + str(WXList_count) + ".xls"
            print(WXList[(c - 1) * ROW:WXList_count])
            print(xls_file_name)
            sheet1.range('A2').options(transpose=True).value = WXList[(c - 1) * ROW:WXList_count]
        else:
            xls_file_name = xls_path2 + str((c - 1) * ROW) + "-" + str(c * ROW) + ".xls"
            print(WXList[(c - 1) * ROW:c * ROW])
            print(xls_file_name)
            sheet1.range('A2').options(transpose=True).value = WXList[(c - 1) * ROW:c * ROW]
        wb.save(xls_file_name)
        wb.close()
        app.quit()