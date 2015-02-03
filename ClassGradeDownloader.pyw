# -*- coding: UTF-8 -*-

__author__ = 'FinalTheory'

from CommonGUI import *
from tkMessageBox import showinfo, showerror, askyesno


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
        self.flag.set('1')

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
        # 如果下载的是原始成绩，才对其进行绘图分析
        if self.flag.get() == '1':
            # 首先尝试导入所有依赖库
            try:
                import os
                import numpy as np
                import BeautifulSoup
                import matplotlib.pyplot as plt
                from matplotlib.font_manager import FontProperties
            except:
                showerror(u'系统错误', u'未能导入成绩绘图分析所需要的库，请自行打开文件查看。')
                return

            # 首先取出所有分值字符串
            scores = []
            soup = BeautifulSoup.BeautifulSoup(data)
            for idx, item in enumerate(soup.findAll('td', {'style': "text-align: right;"})):
                if (idx + 1) % 5 == 0:
                    scores.append(item.renderContents())

            # 然后尝试将字符串转换为浮点数
            # 若失败则说明成绩不以数值表示，也直接返回
            try:
                scores = map(float, scores)
            except ValueError:
                return

            scores = np.array(scores)

            # 如果成绩数据为空，说明成绩还没出来
            if len(scores) == 0:
                # 此时提示是否删除已经下载的文件
                if askyesno(u'系统错误', u'没有读取到有效成绩，可能教师尚未登入分数。\n是否删除已下载的文件？'):
                    os.remove(filename)
            else:
                # 显示统计信息，询问是否绘图
                if askyesno(u'统计信息',
                            u'选课：%d人\n通过：%d人\n未通过：%d人\n其中旷考/缓考：%d人\n是否显示详细分布图？'
                            % (len(scores), sum(scores >= 60), sum(scores < 60), sum(scores == 0))):

                    # 设置中文字体
                    font = FontProperties(fname=r"c:\windows\fonts\msyh.ttc", size=15)
                    plt.figure()
                    x = np.linspace(0, 100, 11)
                    n, bins, patches = plt.hist(scores,
                                                x,
                                                normed=0,
                                                histtype='bar',
                                                rwidth=0.9)
                    plt.title(u'成绩分布示意图', fontproperties=font)
                    plt.xlabel(u'成绩区间', fontproperties=font)
                    plt.ylabel(u'人数', fontproperties=font)
                    plt.xticks(x)
                    plt.show()

if __name__ == '__main__':
    GUI(310, 300).mainloop()