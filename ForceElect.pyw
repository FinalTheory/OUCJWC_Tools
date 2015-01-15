# -*- coding: UTF-8 -*-

__author__ = 'FinalTheory'

import json
from CommonGUI import *
from copy import deepcopy
from tkMessageBox import showinfo, showerror


class Worker(CommonWorker):
    def __init__(self):
        CommonWorker.__init__(self)
        # 进行业务处理时真正的目标地址
        self.targetURL = 'http://jwgl.ouc.edu.cn/jw/common/saveElectiveCourse.action'
        self.targetRefURL = 'http://jwgl.ouc.edu.cn/student/wsxk.axkhksxk.html?menucode=JW130410'
        # 这个页面返回与加密处理相关的信息
        self.encryptURL = 'http://jwgl.ouc.edu.cn/custom/js/SetKingoEncypt.jsp'
        # 下面是普通js脚本的地址
        self.jsList = [
            'http://jwgl.ouc.edu.cn/custom/js/jkingo.des.js',
            'http://jwgl.ouc.edu.cn/custom/js/base64.js',
            'http://jwgl.ouc.edu.cn/custom/js/md5.js'
        ]
        # 下面这些地址用于获取查询信息
        self.queryURL = [
            'http://jwgl.ouc.edu.cn/jw/common/getCourseInfoByXkh.action',
            'http://jwgl.ouc.edu.cn/jw/common/getSelectLessonScoreKcsInfo.action',
            'http://jwgl.ouc.edu.cn/jw/common/getSelectLessonPointsInfo.action',
            'http://jwgl.ouc.edu.cn/jw/common/getStuGradeSpeciatyInfo.action',

        ]
        self.queryHeaderURL = [
            'http://jwgl.ouc.edu.cn/student/wsxk.axkhksxk.html?menucode=JW130410',
            'http://jwgl.ouc.edu.cn/student/wsxk.axkhksxk.html?menucode=JW130410',
            'http://jwgl.ouc.edu.cn/student/wsxk.axkhksxk.html?menucode=JW130410',
            'http://jwgl.ouc.edu.cn/student/wsxk.axkhksxk.html?menucode=JW130410',
        ]

    def QueryClass(self, xkh, Year, Term, xh, weight, is_cx, is_buybook, is_yxtj='1'):

        postData = {
            'xkh': xkh,
            'xn': Year,
            'xq': Term,
            'xq_m': Term,
            'xh': xh,
            'hidOption': 'qry'
        }

        allData = deepcopy(postData)
        for queryURL, queryHeaderURL in zip(self.queryURL, self.queryHeaderURL):
            req = urllib2.Request(url=queryURL)
            req.add_header('Referer', queryHeaderURL)
            response = self.opener.open(req, urllib.urlencode(postData))
            data = json.loads(response.read())
            # 把所有数据都转换成字典
            if type(data) == list and len(data) > 0:
                data = data[0]
            elif type(data) == dict:
                data = json.loads(data['result'])
            # 然后保存到一起去
            if type(data) == dict:
                for key in data.keys():
                    allData[key] = data[key]

        try:
            xkData = {
                # 最重要：操作类型
                # 如果设置成4，可以任意删除课程！
                'xktype': '2',
                'is_cx': is_cx,
                'is_buy_book': is_buybook,
                'is_yxtj': is_yxtj,
                'text_weight': weight,
                'xk_points': weight,
                'kkxq': allData['xqmc'],
                'kcrkjs': allData['xm'],
                'xxyx': allData['rs'],
                'skdd': allData['fjbh'],
                'kcxf': allData['xf'],
                'skbjdm': xkh,
                'xkdetails': ','.join([allData['kcdm'], xkh, allData['xf'], is_buybook, is_cx, weight]),
                # 'ck_skbtj': 'on',
                # 'ck_gmjc': 'on',
            }
        except KeyError:
            showerror(u'杯具啊', u'没有检索到课程相关信息，请检查你所选择的学年学期是否正确！')
            return

        self.doPostData(dict(xkData.items() + allData.items()))

    def doPostData(self, postData):
        params = urllib.urlencode(postData)
        try:
            # 模拟执行js，得到加密参数
            from PyV8 import JSContext
            ctxt = JSContext()
            ctxt.enter()
            for url in self.jsList:
                data = self.opener.open(url).read()
                ctxt.eval(data)
            jsContents = self.opener.open(self.encryptURL).read().replace('document', '//document')
            ctxt.eval(jsContents)
        except:
            showerror(u'杯具啊', u'程序运行所需要的js脚本导入失败！')
            return

        _params = ctxt.eval('getEncParams("%s");' % params)
        req = urllib2.Request(url=self.targetURL)
        req.add_header('Referer', self.targetRefURL)
        response = self.opener.open(req, _params)
        result = json.loads(response.read())
        showinfo(u'操作完成', u'返回结果：' + result['message'])


class GUI(CommonGUI):
    def __init__(self, width, height):
        CommonGUI.__init__(self, width, height)
        self.Downloader = Worker()
        self.skbjdm = tk.StringVar()
        self.is_cx = tk.StringVar()
        self.is_cx.set('0')
        self.is_buy_book = tk.StringVar()
        self.is_buy_book.set('0')
        self.weight = tk.StringVar()
        self.weight.set('0')

        #*******************************生成额外功能*******************************
        start_Button = tk.Button(self.frame, text=u'强行选课', width=20, height=2, state='disabled',
                                      command=self.ForceElect, font=('YaHei', 17))
        start_Button.grid(row=9, column=0, columnspan=3)
        self.disabledList.append(start_Button)
        tk.Label(self.frame, text=u'选课号').grid(row=6, column=0, sticky=tk.W)
        tk.Entry(self.frame, textvariable=self.skbjdm).grid(row=6, column=1, sticky=tk.W)

        tk.Label(self.frame, text=u'权重分：').grid(row=7, column=2)
        tk.Entry(self.frame, textvariable=self.weight, width=6).grid(row=8, column=2)

        tk.Label(self.frame, text=u'是否重修').grid(row=7, column=0, sticky=tk.W)
        tk.Radiobutton(self.frame, text=u"重修", variable=self.is_cx, value='1').grid(row=7, column=1, sticky=tk.W)
        tk.Radiobutton(self.frame, text=u"不重修", variable=self.is_cx, value='0').grid(row=7, column=1, sticky=tk.E)

        tk.Label(self.frame, text=u'是否买书').grid(row=8, column=0, sticky=tk.W)
        tk.Radiobutton(self.frame, text=u"买书", variable=self.is_buy_book, value='1').grid(row=8, column=1, sticky=tk.W)
        tk.Radiobutton(self.frame, text=u"不买书", variable=self.is_buy_book, value='0').grid(row=8, column=1, sticky=tk.E)

        #*****************************在这里刷新验证码!****************************
        # self.RefreshImg()
        self.frame.grid()

    def ForceElect(self):
        # 对输入进行简单检查
        try:
            int(self.weight.get())
        except:
            showerror(u'严重错误', u'输入的选课权重分必须为整数！')
            return

        if len(self.skbjdm.get()) == 0:
            showerror(u'严重错误', u'必须输入选课号！')
            return

        self.Downloader.QueryClass(
            self.skbjdm.get(),
            self.yearOption.get(),
            self.termOption.get()[0],
            self.username.get(),
            self.weight.get(),
            self.is_cx.get(),
            self.is_buy_book.get()
        )

if __name__ == '__main__':
    GUI(305, 280).mainloop()
