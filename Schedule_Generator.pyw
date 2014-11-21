# -*- coding: UTF-8 -*-

__author__ = 'FinalTheory'


import csv
import BeautifulSoup
from CommonGUI import *
from base64 import encodestring
from tkFileDialog import askopenfilename


class Worker(CommonWorker):
    def __init__(self):
        CommonWorker.__init__(self)
        self.downloadURL = 'http://jwgl.ouc.edu.cn/student/wsxk.xskcb.jsp?params='

    def download(self, Year, Term, ID):
        params = encodestring("xn=%s&xq=%s&xh=%s" % (Year, Term, ID))
        req = urllib2.Request(url=self.downloadURL + params)
        req.add_header('Referer', 'http://jwgl.ouc.edu.cn/student/xkjg.wdkb.jsp?menucode=JW130416')
        response = self.opener.open(req)
        data = response.read()
        # with open(ID + '.xls', 'w') as fid:
        #     fid.write(data)
        return data


class GUI(CommonGUI):
    def __init__(self, width, height):
        CommonGUI.__init__(self, width, height)

        self.IDMap = {}
        self.Downloader = Worker()
        self.fileName = tk.StringVar()

        #*******************************生成额外功能*******************************
        tk.Label(self.frame, text=u'学号列表').grid(row=5, column=0, sticky=tk.W)
        tk.Entry(self.frame, textvariable=self.fileName).grid(row=5, column=1, sticky=tk.W)
        tk.Button(self.frame, text=u'选择学号文件', command=self.getName).grid(row=5, column=2, sticky=tk.W)

        start_Button = tk.Button(self.frame, text=u'开始抓取', width=20, height=2, state='disabled',
                                      command=self.StartGenerate, font=('YaHei', 17))
        start_Button.grid(row=6, column=0, columnspan=3)
        self.disabledList.append(start_Button)

        #*****************************在这里刷新验证码!****************************
        # self.RefreshImg()
        self.frame.grid()

    def getName(self):
        filename = askopenfilename(title=u'选择有效学号列表')
        if filename:
            self.fileName.set(filename)
            self.IDMap = open(filename, 'r').readlines()
            self.IDMap = filter(lambda x: len(x) > 0,
                                 map(lambda x: x.strip('\n').strip('\r'), self.IDMap))
            self.IDMap = dict(map(lambda x: tuple(x.split()), self.IDMap))
            # print self.IDMap

    def ParseData(self, data):
        def Unit2Status(col_data):
            def Process(item):
                item = str(item)
                return item.replace('<b>', '').replace('</b>', '')
            if 'div_nokb' in col_data:
                return []
            else:
                return map(Process, col_data.findAll('b'))
        result = []
        soup = BeautifulSoup.BeautifulSoup(data)
        rows = soup.findAll('tr')[1:]
        for row in rows:
            cols = row.findAll('td', {'class': "td", 'style': "width:13%;"})
            result.append(map(Unit2Status, cols))
        return result

    def StartGenerate(self):
        AllData = self.DownloadAll()
        UserStatus = map(self.ParseData, AllData)
        csvfile = file(u'成员无课表.csv', 'wb')
        writer = csv.writer(csvfile)
        tableHeader = [u'姓名']
        tableHeader.extend([
            u'周'+ day + cls + u'节课'
            for day in [u'一', u'二', u'三', u'四', u'五', u'六', u'日']
            for cls in ['12', '34', '56', '78', '90']
        ])
        writer.writerow(tableHeader)
        for ID, user in zip(self.IDMap.keys(), UserStatus):
            row_data = [self.IDMap[ID].decode('gb18030').encode('gb18030')]
            for j in range(0, 7):
                for i in range(0, 5):
                    if user[i][j]:
                        row_data.append(' '.join(user[i][j]).decode('utf8').encode('gb18030'))
                    else:
                        row_data.append('')
            writer.writerow(row_data)

        csvfile.close()
        self.PopupWindow(UserStatus)

    def PopupWindow(self, UserStatus):

        def ShowDetail(row, col):
            BusyList = []
            for ID, user in zip(self.IDMap.keys(), UserStatus):
                if user[row-1][col-1]:
                    BusyList.append(self.IDMap[ID].
                                    decode('gb18030').encode('utf8') + ' : ' + ' '.join(user[row-1][col-1]))
            if BusyList: tkMessageBox.showinfo(u"有课人员列表", '\n'.join(BusyList))

        def NumIsBusy(row, col):
            num = 0
            for user in UserStatus:
                if user[row-1][col-1]:
                    num += 1
            return num

        top = tk.Toplevel()
        top.title(u"人员无课明细")

        ColNames = [
            u'时间',
            u'周一',
            u'周二',
            u'周三',
            u'周四',
            u'周五',
            u'周六',
            u'周日'
        ]

        RowNames = [ u'12节', u'34节', u'56节', u'78节', u'90节' ]

        for col in range(0, 8):
            tk.Label(top, text=ColNames[col]).grid(row=0, column=col)

        for row in range(1, 6):
            tk.Label(top, text=RowNames[row-1]).grid(row=row, column=0)

        for row in range(1, 6):
            for col in range(1, 8):
                num = NumIsBusy(row, col)
                tk.Button(top, command=lambda r=row, c=col: ShowDetail(r, c), bg='red' if num > 0 else 'green',
                          width=10, height=2, text=str(num) if num > 0 else '').grid(row=row, column=col)

    def DownloadAll(self):
        if not self.IDMap: self.IDMap = {self.username.get(): u'自己'}
        AllData = []
        for ID in self.IDMap.keys():
            data = self.Downloader.download(self.yearOption.get(), self.termOption.get()[0], ID)
            AllData.append(data)
        return AllData


if __name__ == '__main__':
    GUI(300, 230).mainloop()