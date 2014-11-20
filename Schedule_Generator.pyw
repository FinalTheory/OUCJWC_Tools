# -*- coding: UTF-8 -*-

__author__ = 'FinalTheory'

import csv
import urllib
import urllib2
import hashlib
import datetime
import StringIO
import cookielib
import tkMessageBox
import BeautifulSoup
import Tkinter as tk
from json import loads
from PIL import ImageTk, Image
from base64 import encodestring
from tkFileDialog import askopenfilename


class Worker:
    def __init__(self):
        self.indexURL = 'http://jwgl.ouc.edu.cn/cas/login.action'
        self.authURL = 'http://jwgl.ouc.edu.cn/cas/logon.action'
        self.verifyURL = 'http://jwgl.ouc.edu.cn/cas/genValidateCode?dateTime='
        self.downloadURL = 'http://jwgl.ouc.edu.cn/student/wsxk.xskcb.jsp?params='
        self.Headers = {
            'User-Agent':
            '''Mozilla/5.0 (compatible; MSIE 8.0; Windows NT 6.1; Trident/4.0;
            GTB7.4; InfoPath.2; SV1; .NET CLR 3.3.69573; WOW64; en-US)'''
        }
        cookie = cookielib.CookieJar()
        cookie_handler = urllib2.HTTPCookieProcessor(cookie)
        self.opener = urllib2.build_opener(cookie_handler)
        #根本不需要读取首页，直接POST数据即可
        #fid_index = self.opener.open(urllib2.Request(url=self.indexURL))
        #idx_data = fid_index.read()

    def Login(self, username, password, verify):
        self.postData = {
            'username': username,
            'password': self.md5HashPasswd(password, verify),
            'randnumber': verify.lower(),
            'isPasswordPolicy': '1'
        }
        CurrentPostData = urllib.urlencode(self.postData)
        #增加更多的Headers可以用add_header方法
        req = urllib2.Request(url=self.authURL, data=CurrentPostData, headers=self.Headers)
        response = self.opener.open(req)
        json_data = loads(response.read())
        print json_data
        return json_data['status'] == '200'

    def get_md5_value(self, src):
        myMd5 = hashlib.md5()
        myMd5.update(src)
        myMd5_Digest = myMd5.hexdigest()
        return myMd5_Digest

    def md5HashPasswd(self, passwd, verify):
        return self.get_md5_value(self.get_md5_value(passwd) + self.get_md5_value(verify.lower()))

    def RefreshImg(self):
        date_str = urllib.quote(datetime.datetime.now().
                    strftime(u'%a %b %d %Y %X GMT+0800 (中国标准时间)'), '():+')
        fid_verify = self.opener.open(urllib2.Request(url=self.verifyURL + date_str))
        im_data = fid_verify.read()
        self.im = Image.open(StringIO.StringIO(im_data))

    def download(self, Year, Term, ID):
        params = encodestring("xn=%s&xq=%s&xh=%s" % (Year, Term, ID))
        req = urllib2.Request(url=self.downloadURL + params)
        req.add_header('Referer', 'http://jwgl.ouc.edu.cn/student/xkjg.wdkb.jsp?menucode=JW130416')
        response = self.opener.open(req)
        data = response.read()
        return data
        # with open(ID + '.xls', 'w') as fid:
        #     fid.write(data)


class GUI(tk.Tk):
    def __init__(self, width, height):
        tk.Tk.__init__(self)

        self.username = tk.StringVar()
        self.password = tk.StringVar()
        self.verify = tk.StringVar()
        self.yearOption = tk.StringVar()
        self.termOption = tk.StringVar()
        self.fileName = tk.StringVar()
        self.IDMap = {}
        self.Downloader = Worker()

        #设置界面
        screenwidth = self.winfo_screenwidth()
        screenheight = self.winfo_screenheight()

        size = '%dx%d+%d+%d' % (width, height, (screenwidth - width)/2, (screenheight - height)/2)
        self.geometry(size)

        self.title(u'用户登录')
        self.protocol("WM_DELETE_WINDOW", lambda: (self.quit(), self.destroy()))
        self.frame = tk.Frame(self)
        tk.Label(self.frame, text=u'用户名').grid(row=0, column=0, sticky=tk.W)
        tk.Label(self.frame, text=u'密码').grid(row=1, column=0, sticky=tk.W)
        tk.Label(self.frame, text=u'验证码').grid(row=2, column=0, sticky=tk.W)
        tk.Entry(self.frame, textvariable=self.username).grid(row=0, column=1)
        tk.Entry(self.frame, textvariable=self.password, show='*').grid(row=1, column=1)
        tk.Entry(self.frame, textvariable=self.verify).grid(row=2, column=1)

        self.login_Button = tk.Button(self.frame, text=u'登录', width=12, height=3,
                                      command=self.Login, font=('YaHei', 11))
        self.login_Button.grid(row=0, column=2, rowspan=3)
        self.verify_IMG = tk.Label(self.frame)
        self.verify_IMG.grid(row=3, column=2, rowspan=2)
        self.verify_IMG.bind('<Button-1>', self.RefreshImg)
        self.RefreshImg()

        tk.Label(self.frame, text=u'学年').grid(row=3, column=0, sticky=tk.W)
        cur_year = datetime.datetime.now().year
        self.yearOption.set(str(cur_year))
        OPTIONS = map(str, range(cur_year-2, cur_year+3))
        year_List = apply(tk.OptionMenu, (self.frame, self.yearOption) + tuple(OPTIONS))
        year_List.grid(row=3, column=1, sticky=tk.W)

        tk.Label(self.frame, text=u'学期').grid(row=4, column=0, sticky=tk.W)
        self.termOption.set(u'1（秋季）')
        tk.OptionMenu(self.frame, self.termOption, u'0（夏季）', u'1（秋季）', u'2（春季）').grid(row=4, column=1, sticky=tk.W)

        tk.Label(self.frame, text=u'学号列表').grid(row=5, column=0, sticky=tk.W)
        tk.Entry(self.frame, textvariable=self.fileName).grid(row=5, column=1, sticky=tk.W)
        tk.Button(self.frame, text=u'选择学号文件', command=self.getName).grid(row=5, column=2, sticky=tk.W)

        self.start_Button = tk.Button(self.frame, text=u'开始抓取', width=20, height=2, state='disabled',
                                      command=self.StartGenerate, font=('YaHei', 17))
        self.start_Button.grid(row=6, column=0, columnspan=3)

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

    def RefreshImg(self, *args):
        self.Downloader.RefreshImg()
        self.img = ImageTk.PhotoImage(self.Downloader.im)
        self.verify_IMG.config(image=self.img)

    def Login(self):
        if self.Downloader.Login(self.username.get(), self.password.get(), self.verify.get()):
            self.start_Button.config(state='normal')
            self.login_Button.config(state='disabled')
            tkMessageBox.showinfo(u"恭喜你！", '登录成功！')
            if not self.IDMap:
                self.IDMap = {self.username.get(): u'自己'}
        else:
            tkMessageBox.showerror(u"悲剧啊！", '登录失败！')

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
                                    decode('gb18030').encode('utf8') + ': ' + ' '.join(user[row-1][col-1]))
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
        AllData = []
        for ID in self.IDMap.keys():
            data = self.Downloader.download(self.yearOption.get(), self.termOption.get()[0], ID)
            AllData.append(data)
        return AllData

if __name__ == '__main__':
    GUI(300, 230).mainloop()