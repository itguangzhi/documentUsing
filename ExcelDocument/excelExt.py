

import xlrd
import xlwt

fileDIR = r'E:\chromDownload\6moth.xlsx'
def sheetinfo(excels, sheetx ,nowx):
    num9values = excels.sheet_by_index(sheetx).row_values(nowx)
    days = 1
    for dates in num9values:
        # print(str(days) + '日')
        if dates != '':
            daka = str(dates).split('\n')

            if len(daka) <= 2:
                print('打卡异常')
            moring = daka[0]
            night = daka[-2]
            print(str(days) + '日 '+'早: ' + moring + ' 晚: ' + night)
            days += 1
        else:
            days += 1
            continue

excel = xlrd.open_workbook(fileDIR)
numrows = excel.sheet_by_index(0).nrows

for i in range(9, numrows):
    print(i)
    print(sheetinfo(excel, 0, i))



