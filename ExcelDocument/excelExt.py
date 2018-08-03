#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''
                            _ooOoo_
                           o8888888o
                           88" . "88
                          (|  -_-  |)
                           O\  =  /O
                        ____/`---'\____
                      .   ' \\| |// `.
                       / \\||| : |||// \
                     / _||||| -:- |||||- \
                       | | \\\ - /// | |
                     | \_| ''\---/'' | |
                      \ .-\__ `-` ___/-. /
                   ___`. .' /--.--\ `. . __
                ."" '< `.___\_<|>_/___.' >'"".
               | | : `- \`.;`\ _ /`;.`/ - ` : | |
                 \ \ `-. \_ __\ /__ _/ .-` / /
         ======`-.____`-.___\_____/___.-`____.-'======
                            `=---='

         .............................................
                  佛祖镇楼                  BUG辟易
          佛曰:
                  写字楼里写字间，写字间里程序员；
                  程序人员写程序，又拿程序换酒钱。
                  酒醒只在网上坐，酒醉还来网下眠；
                  酒醉酒醒日复日，网上网下年复年。
                  但愿老死电脑间，不愿鞠躬老板前；
                  奔驰宝马贵者趣，公交自行程序员。
                  别人笑我忒疯癫，我笑自己命太贱；
                  不见满街漂亮妹，哪个归得程序员？
'''
# @File  : ${NAME}.py
# @Author: huguangzhi
# @ContactEmail : huguangzhi@ucsdigital.com.com
# @ContactPhone : 13121961510
# @Date  : ${DATE} - ${TIME}
# @Desc  :

import xlrd
import xlwt




class KPI_earn():
    def sheetinfo(excels, sheetx, nowx):
        libR = {}
        libR['没有打卡日期'] = []
        libR['打卡异常日期'] = []
        rownumvalues = excels.sheet_by_index(sheetx).row_values(nowx)
        days = 1
        for dates in rownumvalues:
            # print(str(days) + '日')
            libR[str(days)] = {}
            if dates != '':
                daka = str(dates).split('\n')
                moring = daka[0]
                night = ''
                if len(daka) <= 2:
                    libR['打卡异常日期'].append(str(days))
                    # print('打卡异常日期')
                    if moring > '13:00':
                        moring = '-'
                        night = daka[0]
                    elif moring < '13:00':
                        night = '-'
                else:
                    night = daka[-2]

                libR[str(days)]['am'] = moring
                libR[str(days)]['pm'] = night
                # print(str(days) + '日 ' + '早: ' + moring + ' 晚: ' + night)
                days += 1
            else:
                # print(str(days) + '日没打卡~')
                libR['没有打卡日期'].append(str(days))
                libR[str(days)]['am'] = '-'
                libR[str(days)]['pm'] = '-'
                days += 1
                continue
        return libR

    def getdaka(self, filePath):
        namelist = {}
        excel = xlrd.open_workbook(filePath)
        numrows = excel.sheet_by_index(0).nrows

        for i in range(9, numrows + 1):
            if i % 2 == 1:
                name = excel.sheet_by_index(0).row_values(i - 1)[10]
                namelist[name] = KPI_earn.sheetinfo(excel, 0, i)
            else:
                continue

        return namelist


fileDIR = r'../file/7月打卡记录.xlsx'
ucslist = KPI_earn.getdaka(KPI_earn, fileDIR)
for i in ucslist:
    print(i)
    print(ucslist[i])

