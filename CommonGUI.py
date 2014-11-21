# -*- coding: UTF-8 -*-

__author__ = 'FinalTheory'

import os
import urllib
import urllib2
import hashlib
import datetime
import StringIO
import cookielib
import tkMessageBox
import Tkinter as tk
from json import loads
from PIL import ImageTk, Image, ImageFont, ImageDraw


class CommonWorker:

    #********************************外部可调用方法********************************

    def __init__(self):

        #*******************************全局共用数据*******************************
        self.indexURL = 'http://jwgl.ouc.edu.cn/cas/login.action'
        self.authURL = 'http://jwgl.ouc.edu.cn/cas/logon.action'
        self.verifyURL = 'http://jwgl.ouc.edu.cn/cas/genValidateCode?dateTime='
        self.Headers = {
            'User-Agent':
            '''Mozilla/5.0 (compatible; MSIE 8.0; Windows NT 6.1; Trident/4.0;
            GTB7.4; InfoPath.2; SV1; .NET CLR 3.3.69573; WOW64; en-US)'''
        }
        cookie = cookielib.CookieJar()
        cookie_handler = urllib2.HTTPCookieProcessor(cookie)
        self.opener = urllib2.build_opener(cookie_handler)
        # 登录时根本不需要读取首页，直接POST数据即可
        # fid_index = self.opener.open(urllib2.Request(url=self.indexURL))
        # idx_data = fid_index.read()

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
        # print CurrentPostData
        return json_data['status'] == '200'

    def RefreshImg(self):
        date_str = urllib.quote(datetime.datetime.now().
                    strftime(u'%a %b %d %Y %X GMT+0800 (中国标准时间)'), '():+')
        fid_verify = self.opener.open(urllib2.Request(url=self.verifyURL + date_str))
        im_data = fid_verify.read()
        self.im = Image.open(StringIO.StringIO(im_data))

    #********************************内部工具函数*********************************

    def get_md5_value(self, src):
        myMd5 = hashlib.md5()
        myMd5.update(src)
        myMd5_Digest = myMd5.hexdigest()
        return myMd5_Digest

    def md5HashPasswd(self, passwd, verify):
        return self.get_md5_value(self.get_md5_value(passwd) + self.get_md5_value(verify.lower()))


class CommonGUI(tk.Tk):
    def __init__(self, width, height):

        #******************************初始化基本变量******************************
        tk.Tk.__init__(self)
        self.username = tk.StringVar()
        self.password = tk.StringVar()
        self.verify = tk.StringVar()
        self.yearOption = tk.StringVar()
        self.termOption = tk.StringVar()
        self.Downloader = CommonWorker()
        # 这个列表记录了未登录时被disabled的控件
        self.disabledList = []
        #由于初始状态下可以不输入验证码，故直接展示提示图片
        self.img = Image.new('RGB', (70, 30), (0, 255, 0))
        font = ImageFont.truetype('msyh.ttc', 17)
        ImageDraw.Draw(self.img).text((1, 4), u'直接登录', font=font, fill=(0, 0, 0))
        self.img = ImageTk.PhotoImage(self.img)

        #*******************************仅供调试使用*******************************
        if os.path.isfile('passwd.txt'):
            dat = map(lambda x: x.rstrip('\n').rstrip('\r'), open('passwd.txt', "r").readlines())
            self.username.set(dat[0])
            self.password.set(dat[1])

        #*******************************生成基本界面*******************************
        self.InitSize(width, height)
        self.InitGUI()

    def InitSize(self, width, height):
        # 设置界面大大小
        screenwidth = self.winfo_screenwidth()
        screenheight = self.winfo_screenheight()

        size = '%dx%d+%d+%d' % (width, height, (screenwidth - width)/2, (screenheight - height)/2)
        self.geometry(size)

    def InitGUI(self):
        # 设置基本元素
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
        self.verify_IMG = tk.Label(self.frame,
                                   image=self.img)
        self.verify_IMG.grid(row=3, column=2, rowspan=2)
        self.verify_IMG.bind('<Button-1>', self.RefreshImg)

        tk.Label(self.frame, text=u'学年').grid(row=3, column=0, sticky=tk.W)
        cur_year = datetime.datetime.now().year
        self.yearOption.set(str(cur_year))
        OPTIONS = map(str, range(cur_year-2, cur_year+3))
        year_List = apply(tk.OptionMenu, (self.frame, self.yearOption) + tuple(OPTIONS))
        year_List.grid(row=3, column=1, sticky=tk.W)

        tk.Label(self.frame, text=u'学期').grid(row=4, column=0, sticky=tk.W)
        self.termOption.set(u'1（秋季）')
        tk.OptionMenu(self.frame, self.termOption, u'0（夏季）', u'1（秋季）', u'2（春季）').\
            grid(row=4, column=1, sticky=tk.W)

        self.frame.grid()

    def RefreshImg(self, *args):
        self.Downloader.RefreshImg()
        self.img = ImageTk.PhotoImage(self.Downloader.im)
        self.verify_IMG.config(image=self.img)

    def Login(self):
        if self.Downloader.Login(self.username.get(), self.password.get(), self.verify.get()):
            self.login_Button.config(state='disabled')
            for item in self.disabledList:
                item.config(state='normal')
            tkMessageBox.showinfo(u"恭喜你！", '登录成功！')
        else:
            tkMessageBox.showerror(u"悲剧啊！", '登录失败！')
            self.RefreshImg()


if __name__ == '__main__':
    CommonGUI(300, 140).mainloop()