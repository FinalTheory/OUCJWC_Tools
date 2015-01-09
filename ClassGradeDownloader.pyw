# -*- coding: UTF-8 -*-

__author__ = 'FinalTheory'

import csv
import BeautifulSoup
from CommonGUI import *
from base64 import encodestring
from tkMessageBox import showinfo


class Worker(CommonWorker):
    def __init__(self):
        CommonWorker.__init__(self)
        self.downloadURL = ''

    def download(self, Year, Term, kcdm, skbjdm, flag='2'):
        if flag == '2':
            self.downloadURL = 'http://jwgl.ouc.edu.cn/wjstgdfw/cjlr.ckxscj.fkcaskbjckcj_rptEffecy_data_exp.jsp'
        else:
            self.downloadURL = 'http://jwgl.ouc.edu.cn/wjstgdfw/cjlr.ckxscj.fkcaskbjckcj_rptOrigina_data_exp.jsp'
        params = {
            'xn': Year,
            'xq': Term[0],
            'kcdm': kcdm,
            'skbjdm': skbjdm,
            'flag': flag,
            'pxfs': 'xhsx'
        }
        params2 = {
            'xn': Year,
            'xq_m': Term[0],
            'kcdm': kcdm,
            'skbjdm': skbjdm,
            'flag': flag,
        }
        print params
        req = urllib2.Request(url=self.downloadURL)
        req.add_header('Referer', 'http://jwgl.ouc.edu.cn/wjstgdfw/cjlr.ckxscj.fkcaskbjckcj_rpt.jsp?'
                       + urllib.urlencode(params2))
        response = self.opener.open(req, urllib.urlencode(params))
        data = response.read()
        # with open(ID + '.xls', 'w') as fid:
        #     fid.write(data)
        return data


class GUI(CommonGUI):
    def __init__(self, width, height):
        CommonGUI.__init__(self, width, height)
        self.Downloader = Worker()
        self.skbjdm = tk.StringVar()
        self.kcdm = tk.StringVar()
        self.flag = tk.StringVar()
        self.clsName = tk.StringVar()
        self.flag = tk.StringVar()
        self.flag.set('2')

        #*******************************生成额外功能*******************************
        start_Button = tk.Button(self.frame, text=u'下载成绩表', width=20, height=2, state='disabled',
                                      command=self.StartDownload, font=('YaHei', 17))
        start_Button.grid(row=9, column=0, columnspan=3)
        self.disabledList.append(start_Button)
        tk.Label(self.frame, text=u'课程号').grid(row=5, column=0, sticky=tk.W)
        tk.Entry(self.frame, textvariable=self.kcdm).grid(row=5, column=1, sticky=tk.W)
        tk.Label(self.frame, text=u'选课号').grid(row=6, column=0, sticky=tk.W)
        tk.Entry(self.frame, textvariable=self.skbjdm).grid(row=6, column=1, sticky=tk.W)
        tk.Label(self.frame, text=u'课程名').grid(row=7, column=0, sticky=tk.W)
        tk.Entry(self.frame, textvariable=self.clsName).grid(row=7, column=1, sticky=tk.W)
        tk.Label(self.frame, text=u'成绩类型').grid(row=8, column=0, sticky=tk.W)
        tk.Radiobutton(self.frame, text=u"原始", variable=self.flag, value='1').grid(row=8, column=1, sticky=tk.W)
        tk.Radiobutton(self.frame, text=u"有效", variable=self.flag, value='2').grid(row=8, column=1, sticky=tk.E)

        #*****************************在这里刷新验证码!****************************
        # self.RefreshImg()
        self.frame.grid()

    def StartDownload(self):
        data = self.Downloader.download(self.yearOption.get(), self.termOption.get(),
                                        self.kcdm.get(), self.skbjdm.get(), self.flag.get())
        filename = self.clsName.get() + '.xls'
        with open(filename, 'w') as fid:
            fid.write(data)
        showinfo(u'操作成功',
                 u'全班成绩已经下载完成，重命名为%s。\n使用Office软件打开时可能会弹出警告，属于正常情况，忽略即可。' % filename)

if __name__ == '__main__':
    GUI(310, 300).mainloop()