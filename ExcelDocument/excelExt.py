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
import datetime

import pymysql
import xlrd
import xlwt


class KPI_earn():

    def sheetinfo(self, excels, sheetx, nowx):
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
                        moring = '09:00'
                        night = daka[0]
                    elif moring < '13:00':
                        night = '17:30'
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
                namelist[name] = KPI_earn.sheetinfo(KPI_earn,excel, 0, i)
            else:
                continue

        return namelist

    def overtimeline(base, beijian):
        realdate = datetime.datetime.strptime(beijian, '%H:%M')
        based = datetime.datetime.strptime(base, '%H:%M')
        if based > realdate:
            jiaban = realdate - realdate
        else:
            jiaban = realdate - based
        timeline = str(int(int(jiaban.seconds) / 1800) * 0.5)
        return timeline

    def signINdata(fileyear: str, filemonth: str):
        LT = []

        # 遍历全公司的人名
        for i in ucslist:
            # 声明每个人的一个月的打卡记录
            ucsname = ucslist[i]
            # 遍历这一个月的打卡记录
            for ri in ucsname:
                SQLDATA = {}
                SQLDATA['real_name'] = real_name = str(i)
                try:
                    day = int(ri)
                    if day < 10:
                        strday = '0' + str(day)
                    else:
                        strday = str(day)
                    SQLDATA['M_day'] = filedate = str(str(fileyear) + str(filemonth) + strday)
                    SQLDATA['ISweekday'] = weekdays = datetime.datetime.strptime(filedate, '%Y%m%d').weekday()
                    SQLDATA['real_am'] = am = ucsname[ri]['am']
                    SQLDATA['real_pm'] = pm = ucsname[ri]['pm']

                    if am == '-':
                        SQLDATA['am'] = '09:00'
                    else:
                        SQLDATA['am'] = am
                    if pm == '-':
                        SQLDATA['pm'] = '17:30'
                    else:
                        SQLDATA['pm'] = pm
                    if am == '-' and pm == '-':
                        SQLDATA['pm'] = '-'
                        SQLDATA['am'] = '-'
                        SQLDATA['overtime_line'] = '-'
                    else:
                        try:
                            SQLDATA['overtime_line'] = KPI_earn.overtimeline('18:00', SQLDATA['pm'])
                        except:
                            pass
                    print(str('每人每天：') + ' : ' + str(SQLDATA))
                    LT.append(SQLDATA)
                    # print(LT)
                    print('&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&')
                except:
                    # 如果不能转为数字则为信息内容
                    print(str(ri) + ' : ' + str(ucsname[ri]))
        return LT

    def saveTomysql(saveinfomation):
        conn = pymysql.connect(
            host='localhost',
            user='spiderinc',
            passwd='spiderinc',
            port=3306,
            charset='utf8',
            database='spiderinc'
        )
        cur = conn.cursor()
        for i in saveinfomation:
            # 数据库每一行的信息
            fileds = []
            values = []
            line = i
            for fi in line:
                fileds.append(fi)
                values.append(line[fi])
            keys = str(fileds).replace('[', '').replace(']', '').replace("'", '')
            value = str(values).replace('[', '').replace(']', '')
            sql = 'insert into kpi_signin(' + keys + ')values(' + value + ');'
            cur.execute(sql)
            print('saveing…………')


fileDIR = r'../file/6moth.xlsx'
ucslist = KPI_earn.getdaka(KPI_earn, fileDIR)

saveinfomation = KPI_earn.signINdata('2018','06')
KPI_earn.saveTomysql(saveinfomation)


