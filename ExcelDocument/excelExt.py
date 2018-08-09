#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
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
"""
# @File  : ${NAME}.py
# @Author: huguangzhi
# @ContactEmail : huguangzhi@ucsdigital.com.com
# @ContactPhone : 13121961510
# @Date  : ${DATE} - ${TIME}
# @Desc  :
import datetime
import tkinter as tk

from Util import Properties
from xlutils.filter import process, XLRDReader, XLWTWriter
import pymysql
import xlrd
from xlutils.copy import copy


class KPI_earn():
    # 数据返回的当月打卡信息，包括未打卡时间和打卡异常时间，以及每天打卡的时间
    def sheetinfo(self, excels, sheetx, nowx):
        libR = {}
        deptvalue = excels.sheet_by_index(sheetx).row_values(nowx - 1)[20]
        libR['部门'] = deptvalue
        libR['没有打卡日期'] = []
        libR['打卡异常日期'] = []
        libR['正常打卡加班日期'] = []

        rownumvalues = excels.sheet_by_index(sheetx).row_values(nowx)
        days = 1
        for dates in rownumvalues:

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
                    if night > '18:30':
                        libR['正常打卡加班日期'].append(str(days))

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

        for i in range(7, numrows + 1):
            if i % 2 == 1:
                name = excel.sheet_by_index(0).row_values(i - 1)[10]
                namelist[name] = KPI_earn.sheetinfo(KPI_earn,excel, 0, i)
            else:
                continue
        return namelist
    #  计算加班时长
    def overtimeline(base, beijian):
        realdate = datetime.datetime.strptime(beijian, '%H:%M')
        based = datetime.datetime.strptime(base, '%H:%M')
        if based > realdate:
            jiaban = realdate - realdate
        else:
            jiaban = realdate - based
        timeline = str(int(int(jiaban.seconds) / 1800) * 0.5)
        return timeline

    def signINdata(fileyear: str, filemonth: str, ucslist):
        # 这是一个数据清洗，将数据清洗为
        #                   人名，
        #                   打卡日期，
        #                   打卡日期是周几，
        #                   真实早晨打卡时间
        #                   真实晚上打卡时间，
        #                   早晨打卡时间（如果没有打卡，用 - 补全）
        #                   晚上打卡时间（如果没有打卡，用 - 补全）
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

                    if am == '-' or pm == '-' or pm < '18:30':
                        SQLDATA['pm'] = '-'
                        SQLDATA['am'] = '-'
                        SQLDATA['overtime_line'] = '-'
                    else:
                        try:
                            SQLDATA['overtime_line'] = KPI_earn.overtimeline('18:00', SQLDATA['pm'])
                        except:
                            pass
                    # print(str('每人每天：') + ' : ' + str(SQLDATA))
                    LT.append(SQLDATA)
                    # print(LT)
                    # print('&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&')
                except:
                    # 如果不能转为数字则为信息内容
                    # print(str(ri) + ' : ' + str(ucsname[ri]))
                    pass
        return LT
    # 获取加班信息
    def overinfo(year: str, month: str, ucslist):
        '''定义变量
            @name 人名，世纪美映所有参与打卡的员工
            @nameinfo 每个人的当月打卡的所有记录信息
            @oneinfo 每个人的每一项的信息内容,包括打卡异常的日期，没有打卡的日期以及每天打卡的信息
        '''
        for name in ucslist:
            print('正在处理' + str(name) + '的打卡记录')
            nameinfo = ucslist[name]
            for oneinfo in nameinfo:
                if oneinfo == '没有打卡日期':
                    signNone = '%s年%s月 '%(year,month)+name + ' 共计有%d天没有打卡' % len(nameinfo[oneinfo])
                    print(signNone)
                    if len(nameinfo[oneinfo]) != 0:
                        signNoneList = '有这几天没有打卡：' + str(nameinfo[oneinfo])
                        print(signNoneList)
                    continue

                elif oneinfo == '打卡异常日期':
                    signErr = name + ' 共计有%d天打卡异常' % len(nameinfo[oneinfo])
                    print(signErr)
                    if len(nameinfo[oneinfo]) != 0:
                        signErrList = '有这几天打卡异常：' + str(nameinfo[oneinfo])
                        print(signErrList)
                        print('这些天的打卡时间如下：')
                        for ERR in nameinfo[oneinfo]:
                            print(nameinfo[ERR])
                    continue

                elif oneinfo == '正常打卡加班日期':
                    signOver = name + ' 共计有%d天加班' % len(nameinfo[oneinfo])
                    print(signOver)
                    if len(nameinfo[oneinfo]) != 0:
                        signErrList = '有这几天打卡异常：' + str(nameinfo[oneinfo])
                        print('这几天加班时间如下：')
                        for ERR in nameinfo[oneinfo]:
                            print(nameinfo[ERR])
                    continue

                else:
                    workingDAY = nameinfo[oneinfo]
                    am = workingDAY['am']
                    pm = workingDAY['pm']
                    if am == '-' or pm == '-' or pm < '18:30':
                        continue
                    else:
                        try:
                            timelines = KPI_earn.overtimeline('18:00', pm)
                        except:
                            print('计算加班时长 ERROR')

            print('-----------------分割线---------------------')
    # 获取存储Excel的加班信息
    def getOverInfoToExcel(self, year: str, month: str, ucslist, overbegin='18:00'):
        """
            {name:{date:{begin:'',end:'',timeline:''},
                   date:begin:'',end:'',timeline:'',
                   date:begin:'',end:'',timeline:''……},
             name:{date:begin:'',end:'',timeline:'',
                   date:begin:'',end:'',timeline:'',
                   date:begin:'',end:'',timeline:''……},
             name:{date:begin:'',end:'',timeline:'',
                   date:begin:'',end:'',timeline:'',
                   date:begin:'',end:'',timeline:''……},}
            """
        overlist = {}
        for name in ucslist:
            overlist[name] = {}
            for D in ucslist[name]['正常打卡加班日期']:
                if int(D) < 10:
                    DA = str(year) + '-' + str(month) + '-0' + D
                else:
                    DA = str(year) + '-' + str(month) + '-' + D
                overlist[name][DA] = {}
                overlist[name][DA]['begin'] = ucslist[name][D]['am']
                overlist[name][DA]['end'] = ucslist[name][D]['pm']
                overlist[name][DA]['timeline'] = KPI_earn.overtimeline(overbegin, ucslist[name][D]['pm'])

        return overlist


class SaveData():

    def saveTomysql(sql):
        PropertiesFile = r'../filename.properties'
        hostname = Properties(PropertiesFile).getProperties()['mysql']['local']['host']
        database = Properties(PropertiesFile).getProperties()['mysql']['local']['database']
        port = int(Properties(PropertiesFile).getProperties()['mysql']['local']['port'])
        username = Properties(PropertiesFile).getProperties()['mysql']['local']['username']
        passwd = Properties(PropertiesFile).getProperties()['mysql']['local']['passwd']
        charset = Properties(PropertiesFile).getProperties()['mysql']['local']['charset']
        conn = pymysql.connect(host=hostname,
                               user=username,
                               passwd=passwd,
                               port=port,
                               db=database,
                               charset=charset
                               )
        try:
            cur = conn.cursor()
        except Exception as e:
            print(e)
            print('-------------连接数据库失败-------------')
        # 执行sql
        else:
            try:
                # print(sql)
                # mysql执行sql语句
                cur.execute(sql)
            except Exception as e:
                print(e)
                print('---SQL语法错误，执行失败---' + sql)
            else:
                conn.commit()
                print('Keep Going ……')

    def builder(savelib):
        DATA = []
        for i in savelib:
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
            DATA.append(sql)
        return DATA

    def builder2(savelib):
        DATA = []
        for i in savelib:
            # 数据库每一行的信息
            fileds = []
            values = []

            line = i
            for fi in line:
                fileds.append(fi)
                values.append(line[fi])
            keys = str(fileds).replace('[', '').replace(']', '').replace("'", '')
            value = str(values).replace('[', '(').replace(']', ')')
            DATA.append(value)
            DATAS = str(DATA).replace('[', '').replace(']', '').replace('"', '')
            sql = 'insert into kpi_signin(' + keys + ')values ' + DATAS + ' ;'
            DATA.append(sql)
        return DATA

    # 加班信息写入excel表格
    def saveOverToExcel(self, file, name, info, begin = '18:00', excelanme = '加班统计表'):

        tablefile = xlrd.open_workbook(file, formatting_info=True, on_demand=True)
        excel, s = SaveData.copy2(tablefile)
        rbs = tablefile.get_sheet(0)
        styles = s[rbs.cell_xf_index(1, 5)] # 引用模板的单元格格式
        styles1 = s[rbs.cell_xf_index(4, 0)]
        styles2 = s[rbs.cell_xf_index(4, 1)]
        styles3 = s[rbs.cell_xf_index(4, 2)]
        styles4 = s[rbs.cell_xf_index(4, 3)]

        # excel = copy(tablefile)  # 用xlutils提供的copy方法将xlrd的对象转化为xlwt的对象
        table = excel.get_sheet(0)
        tablefile.release_resources()  # 关闭模板文件
        table.write(1, 5, name, styles)
        row = 3
        for save in info:
            table.write(int(row), 0, save, styles1)
            # table.write(row, 1, info[save]['begin'], styles2)
            table.write(row, 1, begin, styles2)
            table.write(row, 2, info[save]['end'], styles3)
            timeline = float(info[save]['timeline'])
            table.write(row, 3, timeline, styles4)
            # table.write(row, 3, info[save]['timeline'], styles4)

            row += 1
        excel.save('../exec/'+str(name)+'-'+excelanme+'.xls')

    def copy2(wb):
        w = XLWTWriter()
        process(XLRDReader(wb, '统计表.xls'), w)
        return w.output[0][1], w.style_list

class APP:
    def __init__(self, master):
        frame = tk.Frame(master)
        frame.pack()

        self.button_1 = tk.Button(frame,
                                  text='按钮1',
                                  bg='black',
                                  fg='whilt',
                                  command=self.exportExcel)
        self.button_1.pack()



    def exportExcel(self):
        fileDIR = r'../file/6moth.xlsx'
        modlefile = r'../file/统计表.xls'
        year = '2018'
        month = '06'

        ucslist = KPI_earn.getdaka(KPI_earn, fileDIR)
        print(ucslist)
        overlist = KPI_earn.getOverInfoToExcel(KPI_earn, year, month, ucslist)
        for name in overlist:
            print('正在处理' + name + '的信息')
            overinfo = overlist[name]
            SaveData.saveOverToExcel(SaveData, file=modlefile, name=name, info=overinfo)

if __name__ == '__main__':

    root = tk.Tk()

    app = APP(root)

    root.mainloop()







    #
    # dates = '20180601'
    # KPI_earn.saveOverToExcel(KPI_earn, savefilename, dates, begin='19:32',end='45:65',timeline='2.6')

    # 数据打印
    # KPI_earn.overinfo('2018', '06', ucslist)




    # 开始创建入库方式
#     saveinfomation = KPI_earn.signINdata('2018', '06', ucslist)
#     sqlinfo = SaveData.builder2(saveinfomation)
#     for i in sqlinfo:
#         print(i)
#         # SaveData.saveTomysql(i)

# fileDIR = r'../file/6moth.xlsx'
# print(KPI_earn.getdaka(KPI_earn, fileDIR))

