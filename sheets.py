# coding: utf-8
from openpyxl import load_workbook
from datetime import date, timedelta

wb = load_workbook('greens.xlsx')
ws1 = wb.active

ws2 = wb.copy_worksheet(ws1)
ws3 = wb.copy_worksheet(ws1)
ws4 = wb.copy_worksheet(ws1)
ws5 = wb.copy_worksheet(ws1)

ws1.title = str(date.today())
ws2.title = str(date.today()+timedelta(1))
ws3.title = str(date.today()+timedelta(2))
ws4.title = str(date.today()+timedelta(3))
ws5.title = str(date.today()+timedelta(4))

dicts = {'1': '一', '2': '二', '3': '三', '4': '四', '5': '五', '6': '六', '7': '七'}
ws1['A31'] = '检测日期：%s' % str(date.today()) + '   ' + '星期%s' % dicts[str(date.today().isoweekday())] +\
             '   ' + '检测员：丁杰俊 廖娟'
ws2['A31'] = '检测日期：%s' % str(date.today()+timedelta(1)) + '   ' + '星期%s' % dicts[str((date.today()+timedelta(1)).isoweekday())] +\
             '   ' + '检测员：丁杰俊 廖娟'
ws3['A31'] = '检测日期：%s' % str(date.today()+timedelta(2)) + '   ' + '星期%s' % dicts[str((date.today()+timedelta(2)).isoweekday())] + \
             '   ' + '检测员：丁杰俊 廖娟'
ws4['A31'] = '检测日期：%s' % str(date.today()+timedelta(3)) + '   ' + '星期%s' % dicts[str((date.today()+timedelta(3)).isoweekday())] + \
             '   ' + '检测员：丁杰俊 廖娟'
ws5['A31'] = '检测日期：%s' % str(date.today()+timedelta(4)) + '   ' + '星期%s' % dicts[str((date.today()+timedelta(4)).isoweekday())] + \
             '   ' + '检测员：丁杰俊 廖娟'

wb.save('newgreens.xlsx')

