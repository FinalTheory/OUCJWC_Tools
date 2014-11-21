# -*- coding: UTF-8 -*-

__author__ = 'FinalTheory'


import csv
import BeautifulSoup
from CommonGUI import *
from threading import Thread
from tkFileDialog import askopenfilename, askdirectory
from multiprocessing.dummy import Pool as ThreadPool


def match(name, courseList):
    for item in courseList:
        if item.encode('utf-8') in name:
            return True
    return False


class Student:

    class Info:
        def __init__(self, name, credit, score, status, remark, courseType):
            #记录每门课的课程名、学分、得分、状态、备注，以及是否计入这门课的成绩
            self.isOK = True
            self.name = name
            self.credit = credit
            self.score = score
            self.status = status
            self.remark = remark
            self.courseType = courseType

            #没有分数的课程不计入成绩
            try:
                self.score = float(self.score)
                self.credit = float(self.credit)
            except:
                self.isOK = False

            #只计入初修或者缓考取得的成绩
            if not (r'初修' in self.status or r'缓考' in self.status):
                self.isOK = False

            #初修但是却缓考时成绩可能非零，此时不予计入
            if r'缓考' in self.remark:
                self.isOK = False

            #分数如果小于1则不计入（缓考或者补考）
            if self.score < 1:
                self.isOK = False

    def __init__(self, data, courseList):
        self.userData = []
        self.Included = []
        self.notIncluded = []
        self.courseList = courseList
        self.sumCredit = 0.
        self.sumScore = 0.

        soup = BeautifulSoup.BeautifulSoup(data.decode('gb18030').encode('utf8'))

        results = []

        self.userName = soup.find("table", {"style": "width:100%;border:none;"}).findAll('td')[-1].renderContents().strip('\n')
        self.profession = soup.find("table", {"style": "width:100%;border:none;"}).findAll('td')[1].renderContents().strip('\n')
        self.userName = self.userName.decode('utf8').replace(u"姓名：", "", 1).encode('gb18030')
        self.profession = self.profession.decode('utf8').replace(u"行政班级：", "", 1).encode('gb18030')

        print self.userName

        alltables = soup.findAll("table", {"style": "clear:left;width:199mm;font-size:12px;"})

        for table in alltables:
            rows = table.findAll('tr')
            table_text = []
            for tr in rows:
                row_text = []
                cols = tr.findAll('td')
                for td in cols:
                    text = td.renderContents().strip('\n')
                    row_text.append(text)
                table_text.append(row_text)
            results.append(table_text)

        for table in results:
            for row in table:
                if r'序号' in row[0]:
                    continue
                self.userData.append(self.Info(row[1], row[2], row[6], row[7], row[8], row[3]))

        #列举每一条成绩记录
        #列入计算的成绩需要满足三个条件：
        #1.成绩本身是有效的（初修或者缓考取得）
        #2.这门科目的成绩先前没有被计算过
        #3.这门科目属于必修课程列表
        for info in self.userData:
            if info.isOK and not info.name in self.Included and match(info.name, self.courseList):
                self.Included.append(info.name)
                self.sumCredit += info.credit
                self.sumScore += info.score * info.credit
            else:
                self.notIncluded.append(info)

        self.averScore = self.sumScore / self.sumCredit


class Worker(CommonWorker):

    def __init__(self):
        CommonWorker.__init__(self)
        self.downloadURL = 'http://jwgl.ouc.edu.cn/student/xscj.stuckcj_data_exp.jsp'

    def download(self, Year, Term, ID):
        PostData = {
            'xn': Year,
            'xn1': str(int(Year) + 1),
            'xq': Term,
            'userCode': ID,
            'ysyx': 'yscj',
            'sjxz': 'sjxz1',
            'ysyxS': 'on',
            'sjxzS': 'on'
        }
        req = urllib2.Request(url=self.downloadURL, data=urllib.urlencode(PostData))
        req.add_header('Referer', 'http://jwgl.ouc.edu.cn/student/xscj.stuckcj.jsp')
        response = self.opener.open(req)
        data = response.read()
        # with open(ID + '.xls', 'w') as fid:
        #     fid.write(data)
        return data


#界面逻辑部分
class GUI(CommonGUI):
    def __init__(self, width, height):

        CommonGUI.__init__(self, width, height)

        self.notIncluded = []
        self.courseList = []
        self.courseListName = tk.StringVar()
        self.IDList = []
        self.IDListName = tk.StringVar()
        self.allStudents = []
        self.Downloader = Worker()

        tk.Label(self.frame, text=u'课程列表').grid(row=5, column=0, sticky=tk.W)
        tk.Entry(self.frame, textvariable=self.courseListName).grid(row=5, column=1)
        tk.Button(self.frame, text='...', command=lambda option='course': self.getNameAndFile(option))\
                .grid(row=5, column=2, sticky=tk.W)

        tk.Label(self.frame, text=u'学号列表').grid(row=6, column=0, sticky=tk.W)
        tk.Entry(self.frame, textvariable=self.IDListName).grid(row=6, column=1)
        tk.Button(self.frame, text='...', command=lambda option='ID': self.getNameAndFile(option))\
            .grid(row=6, column=2, sticky=tk.W)
        tk.Button(self.frame, text='自动生成', command=self.AutoGenerate).grid(row=6, column=2, sticky=tk.E)

        start_Button = tk.Button(self.frame, text=u'开始计算', state='disabled', font=('YaHei', 17),
                                 width=20, height=2, command=self.calc)

        start_Button.grid(row=7, column=0, columnspan=3)
        self.disabledList.append(start_Button)

        self.frame.grid()

    def getNameAndFile(self, option):
        filename = askopenfilename(title=u'选择有效课程列表')
        if filename:
            if option == 'course':
                self.courseListName.set(filename)
                self.courseList = open(filename, 'r').readlines()
                self.courseList = filter(lambda x: len(x) > 0,
                                         map(lambda x: x.rstrip('\n').rstrip('\r'), self.courseList))
            elif option == 'ID':
                self.IDListName.set(filename)
                self.IDList = open(filename, 'r').readlines()
                self.IDList = filter(lambda x: len(x) > 0,
                                         map(lambda x: x.rstrip('\n').rstrip('\r'), self.IDList))
            else:
                return

    def DownloadAll(self):
        pool = ThreadPool(16)
        return pool.map(lambda ID, Year=self.yearOption.get(), Term=self.termOption.get():
                        self.Downloader.download(Year, Term, ID), self.IDList)

    def AutoGenerate(self):

        def StartGenerate():
            with open(u'学号列表.txt', 'w') as fid:
                s = self.Start.get()
                e = self.End.get()
                if not s or not e or not self.Prefix.get():
                    tkMessageBox.showerror(u'注意！', u'请输入相应参数！！！')
                    return
                for i in range(int(s), int(e) + 1):
                    Suffix = str(i)
                    while len(Suffix) < max(len(s), len(e)):
                        Suffix = '0' + Suffix
                    fid.write(self.Prefix.get() + Suffix + '\n')
            tkMessageBox.showinfo(u'恭喜你！', u'生成学号列表成功！\n'
            u'如果需要增加新的学号，请直接编辑"学号列表.txt"文件。\n'
            u'完成后，请选择这个文件作为学号列表的输入文件！\n')

        self.Prefix = tk.StringVar()
        self.Start = tk.StringVar()
        self.End = tk.StringVar()
        top = tk.Toplevel()
        top.title(u"学号列表生成器")
        tk.Label(top, text=u'学号前缀').grid(row=0, column=0, sticky=tk.W)
        tk.Label(top, text=u'起始后缀').grid(row=1, column=0, sticky=tk.W)
        tk.Label(top, text=u'结束后缀').grid(row=2, column=0, sticky=tk.W)
        tk.Entry(top, textvariable=self.Prefix).grid(row=0, column=1)
        tk.Entry(top, textvariable=self.Start).grid(row=1, column=1)
        tk.Entry(top, textvariable=self.End).grid(row=2, column=1)
        tk.Button(top, text=u'开始生成', width=20, height=2, command=StartGenerate, font=('YaHei', 17))\
            .grid(row=3, column=0, columnspan=2)
        top.grid()

    def calc(self):

        if not self.IDList or not self.courseList:
            tkMessageBox.showerror(u'杯具啊！', u'请指定学号列表文件以及课程列表文件！')
            return

        self.notIncluded = []
        self.allStudents = []
        AllData = self.DownloadAll()

        tkMessageBox.showinfo(u'恭喜你！', u'所有数据下载完成！\n处理数据可能较慢，并非程序崩溃，请耐心等待。')
        fid = open(u'所有人未被计入成绩的课程列表.txt', 'w')

        for data in AllData:
            cur = Student(data, self.courseList)
            for item in cur.notIncluded:
                if not item.name in self.notIncluded and not match(item.name, self.courseList):
                    self.notIncluded.append(item.name)
                    fid.write(('\t'.join((item.name, str(item.credit), item.courseType)) + '\n')
                              .decode('utf-8').encode('gb18030'))
            self.allStudents.append(cur)

        fid.close()

        #对成绩进行排序
        self.allStudents.sort(key=lambda x: x.averScore, reverse=True)

        fid = open(u'计入的成绩及明细.txt', 'w')
        csvfile = file(u'总成绩排名.csv', 'wb')
        writer = csv.writer(csvfile)
        writer.writerow([u'排名', u'姓名', u'专业年级', u'平均分', u'总学分', u'总学分与成绩之积'])
        index = 1
        for stu in self.allStudents:
            fid.write((stu.userName + '\n'))
            fid.write('*' * 80 + '\n')
            fid.write(u'计入成绩的课程列表：\n')
            fid.write(('\n'.join(stu.Included) + '\n').decode('utf-8').encode('gb18030'))
            fid.write('*' * 80 + '\n')
            fid.write(u'未计入成绩的课程以及明细：\n')
            for course in stu.notIncluded:
                fid.write(('\t'.join((course.name, str(course.credit), str(course.score),
                                      course.status, course.remark)) + '\n').decode('utf-8').encode('gb18030'))
            fid.write('*' * 80 + '\n')
            fid.write('\n\n')
            writer.writerow([str(index),
                             stu.userName,
                             stu.profession,
                             str(stu.averScore),
                             str(stu.sumCredit),
                             str(stu.sumScore)])
            index += 1
        csvfile.close()
        fid.close()

        tkMessageBox.showinfo(u"恭喜你", u"成绩计算完毕！")


if __name__ == "__main__":
    GUI(320, 260).mainloop()