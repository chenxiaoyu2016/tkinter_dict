# -*- coding: UTF-8 -*-
__author__ = 'chenxiaoyu'
from tkinter.messagebox import *
from tkinter.filedialog import *
import os
import xlrd
import xlwt
import xlutils.copy
import random

class application(object):
    def __init__(self):
        self.en = []
        self.pronunc = []
        self.cn = []
        self.length = 0
        self.myindex = 0
        self.type = 0 #功能类型，默认为0，正常取值为1-6
        self.order = []  #存储乱序序列，用于随机考查
        self.times = []  #记录拼错次数
        self.f = NONE    #保存被打开的xls日志文件句柄

    def reinit(self):
        self.en = []
        self.pronunc = []
        self.cn = []
        self.length = 0
        self.myindex = 0
        self.type = 0
        self.order = []
        self.times = []
        self.f = NONE

    def read_data(self,file):
        if(self.f!=NONE):
            self.save_log_file()
        self.reinit()
        textPad.delete(1.0, END)
        try:
            raw_data = xlrd.open_workbook(file)
            table = raw_data.sheets()[0]
            en_raw = table.col_values(0)
            self.en = en_raw[1:len(en_raw)]
            pronunc_raw = table.col_values(1)
            self.pronunc = pronunc_raw[1:len(en_raw)]
            cn_raw = table.col_values(2)
            self.cn = cn_raw[1:len(en_raw)]
            self.length = len(self.en)
            for num in range(self.length):
                self.order.append(num)
        except:
            self.reinit()
            showinfo('错误提示',
                     '您的词典格式不符合要求，请阅读【使用说明】')

        if(self.length>0):
            self.read_record()

    #完成 self.times的初始化
    def read_record(self):
        global myfiles
        if(myfiles.log_exist == 0):
            self.times = [0]*self.length
            self.f = xlwt.Workbook()

        else:
            self.f = xlrd.open_workbook(myfiles.file_path+myfiles.log_name)
            logtable = self.f.sheets()[0]
            self.times = logtable.col_values(1)[1:]

    def print_encn(self):
        for i in range(len(self.cn)):
            print(self.en[i],'\t\t\t\t',self.pronunc[i],'\t\t\t\t',self.cn[i])

    def start_entocn(self):
        self.myindex = 0
        self.type = 1
        textPad.delete(1.0, END)
        textPad.insert(1.0,'总单词数大约：'+str(self.length)+'\n\n')
        # textPad.focus_force()   #使光标在插入的文本之后
        while(self.myindex<self.length and self.cn[self.myindex]==''):    # 不显示无中文翻译的单词
            self.myindex = self.myindex + 1
        if(self.myindex==self.length):
            textPad.insert(CURRENT, '[all done!]' + '\n')
        else:
            textPad.insert(CURRENT,str(self.en[self.myindex])+'\t')

    def entocn_next(self,astr):
        if (self.myindex == self.length):
            textPad.insert(CURRENT, '[all done!]' + '\n')
            textPad.see(END)
            return

        if (astr == loophole):
            textPad.insert(CURRENT,'\n')
            textPad.see(END)
        else:
            textPad.insert(CURRENT, astr + '\n')
            textPad.see(END)
        # with judgement
        if (astr == loophole or astr == self.cn[self.myindex]):
            textPad.insert(CURRENT, "right!", "tag_right")
            textPad.see(END)
            textPad.insert(CURRENT, "\n\n")
            textPad.see(END)
        else:
            textPad.insert(CURRENT, "wrong!", "tag_wrong")
            textPad.see(END)
            textPad.insert(CURRENT, '\t' + str(self.cn[self.myindex]) + '\n\n')
            textPad.see(END)

        self.myindex = self.myindex + 1
        while (self.myindex<self.length and self.cn[self.myindex]==''):  # 不显示无中文翻译的单词
            self.myindex = self.myindex + 1
        if (self.myindex == self.length):
            textPad.insert(CURRENT, '[all done!]' + '\n')
            textPad.see(END)
        else:
            textPad.insert(CURRENT, str(self.en[self.myindex]) + '\t')
            textPad.see(END)

    def start_cntoen(self):
        self.myindex = 0
        self.type = 2
        textPad.delete(1.0, END)
        textPad.insert(1.0,'总单词数大约：'+str(self.length)+'\n\n')
        while(self.myindex<self.length and self.cn[self.myindex]==''):    # 不显示无中文翻译的单词
            self.myindex = self.myindex + 1
        if(self.myindex==self.length):
            textPad.insert(CURRENT, '[all done!]' + '\n')
        else:
            textPad.insert(CURRENT,str(self.cn[self.myindex])+'\t')

    def cntoen_next(self,astr):
        if (self.myindex == self.length):
            textPad.insert(CURRENT, '[all done!]' + '\n')
            textPad.see(END)
            return

        if (astr == loophole):
            textPad.insert(CURRENT,'\n')
            textPad.see(END)
        else:
            textPad.insert(CURRENT, astr + '\n')
            textPad.see(END)
        if(astr == loophole or astr==self.en[self.myindex]):
            # textPad.insert(CURRENT, '[right!]' + '\n\n')  #without tag
            textPad.insert(CURRENT, "right!", "tag_right")
            textPad.see(END)
            textPad.insert(CURRENT,  "\n\n")
            textPad.see(END)

        else:
            self.times[self.myindex] = self.times[self.myindex] + 1
            textPad.insert(CURRENT, "wrong!"+'\t'+'累计错误次数:'+str(int(self.times[self.myindex]))+' ', "tag_wrong")
            textPad.see(END)
            textPad.insert(CURRENT,'\t'+str(self.en[self.myindex])+'\n\n')
            textPad.see(END)
        self.myindex = self.myindex + 1
        while (self.myindex<self.length and self.cn[self.myindex]==''):  # 不显示无中文翻译的单词
            self.myindex = self.myindex + 1
        if (self.myindex == self.length):
            textPad.insert(CURRENT, '[all done!]' + '\n')
            textPad.see(END)
        else:
            textPad.insert(CURRENT, str(self.cn[self.myindex]) + '\t')
            textPad.see(END)

    def start_cntoen_pronunc(self):
        self.myindex = 0
        self.type = 3
        textPad.delete(1.0, END)
        textPad.insert(1.0,'总单词数大约：'+str(self.length)+'\n\n')
        while(self.myindex<self.length and self.cn[self.myindex]==''):    # 不显示无中文翻译的单词
            self.myindex = self.myindex + 1
        if(self.myindex==self.length):
            textPad.insert(CURRENT, '[all done!]' + '\n')
        else:
            textPad.insert(CURRENT,str(self.cn[self.myindex])+'\t')

    def cntoen_pronunc_next(self,astr):
        if (self.myindex == self.length):
            textPad.insert(CURRENT, '[all done!]' + '\n')
            textPad.see(END)
            return

        if (astr == loophole):
            textPad.insert(CURRENT,'\n')
            textPad.see(END)
        else:
            textPad.insert(CURRENT, astr + '\n')
            textPad.see(END)
        if(astr == loophole or astr==self.en[self.myindex]):
            # textPad.insert(CURRENT, '[right!]' + '\n\n')  #without tag
            textPad.insert(CURRENT, "right!", "tag_right")
            textPad.see(END)
            textPad.insert(CURRENT,  "\n\n")
            textPad.see(END)

        else:
            self.times[self.myindex] = self.times[self.myindex] + 1
            textPad.insert(CURRENT, "wrong!" + '\t' + '累计错误次数:' + str(int(self.times[self.myindex])) + ' ', "tag_wrong")
            textPad.see(END)
            textPad.insert(CURRENT,'\t'+str(self.en[self.myindex])+'\t'+str(self.pronunc[self.myindex])+'\n\n')
            textPad.see(END)
        self.myindex = self.myindex + 1
        while (self.myindex<self.length and self.cn[self.myindex]==''):  # 不显示无中文翻译的单词
            self.myindex = self.myindex + 1
        if (self.myindex == self.length):
            textPad.insert(CURRENT, '[all done!]' + '\n')
            textPad.see(END)
        else:
            textPad.insert(CURRENT, str(self.cn[self.myindex]) + '\t')
            textPad.see(END)

    def start_entocn_random(self):
        self.myindex = 0
        self.type = 4
        textPad.delete(1.0, END)
        textPad.insert(1.0,'总单词数大约：'+str(self.length)+'\n\n')
        random.shuffle(self.order)

        while(self.myindex<self.length and self.cn[self.order[self.myindex]]==''):    # 不显示无中文翻译的单词
            self.myindex = self.myindex + 1
        if(self.myindex==self.length):
            textPad.insert(CURRENT, '[all done!]' + '\n')
        else:
            textPad.insert(CURRENT,str(self.en[self.order[self.myindex]])+'\t')

    def entocn_next_random(self,astr):
        if (self.myindex == self.length):
            textPad.insert(CURRENT, '[all done!]' + '\n')
            textPad.see(END)
            return

        if (astr == loophole):
            textPad.insert(CURRENT,'\n')
            textPad.see(END)
        else:
            textPad.insert(CURRENT, astr + '\n')
            textPad.see(END)
        # with judgement
        if (astr == loophole or astr == self.cn[self.order[self.myindex]]):
            textPad.insert(CURRENT, "right!", "tag_right")
            textPad.see(END)
            textPad.insert(CURRENT, "\n\n")
            textPad.see(END)
        else:
            textPad.insert(CURRENT, "wrong!", "tag_wrong")
            textPad.see(END)
            textPad.insert(CURRENT, '\t' + str(self.cn[self.order[self.myindex]]) + '\n\n')
            textPad.see(END)

        self.myindex = self.myindex + 1
        while (self.myindex<self.length and self.cn[self.order[self.myindex]]==''):  # 不显示无中文翻译的单词
            self.myindex = self.myindex + 1
        if (self.myindex == self.length):
            textPad.insert(CURRENT, '[all done!]' + '\n')
            textPad.see(END)
        else:
            textPad.insert(CURRENT, str(self.en[self.order[self.myindex]]) + '\t')
            textPad.see(END)

    def start_cntoen_random(self):
        self.myindex = 0
        self.type = 5
        textPad.delete(1.0, END)
        textPad.insert(1.0,'总单词数大约：'+str(self.length)+'\n\n')
        random.shuffle(self.order)
        while(self.myindex<self.length and self.cn[self.order[self.myindex]]==''):    # 不显示无中文翻译的单词
            self.myindex = self.myindex + 1
        if(self.myindex==self.length):
            textPad.insert(CURRENT, '[all done!]' + '\n')
        else:
            textPad.insert(CURRENT,str(self.cn[self.order[self.myindex]])+'\t')

    def cntoen_next_random(self,astr):
        if (self.myindex == self.length):
            textPad.insert(CURRENT, '[all done!]' + '\n')
            textPad.see(END)
            return

        if (astr == loophole):
            textPad.insert(CURRENT,'\n')
            textPad.see(END)
        else:
            textPad.insert(CURRENT, astr + '\n')
            textPad.see(END)
        if(astr == loophole or astr==self.en[self.order[self.myindex]]):
            # textPad.insert(CURRENT, '[right!]' + '\n\n')  #without tag
            textPad.insert(CURRENT, "right!", "tag_right")
            textPad.see(END)
            textPad.insert(CURRENT,  "\n\n")
            textPad.see(END)

        else:
            self.times[self.order[self.myindex]] = self.times[self.order[self.myindex]] + 1
            textPad.insert(CURRENT, "wrong!" + '\t' + '累计错误次数:' + str(int(self.times[self.order[self.myindex]])) + ' ',"tag_wrong")
            textPad.see(END)
            textPad.insert(CURRENT,'\t'+str(self.en[self.order[self.myindex]])+'\n\n')
            textPad.see(END)
        self.myindex = self.myindex + 1
        while (self.myindex<self.length and self.cn[self.order[self.myindex]]==''):  # 不显示无中文翻译的单词
            self.myindex = self.myindex + 1
        if (self.myindex == self.length):
            textPad.insert(CURRENT, '[all done!]' + '\n')
            textPad.see(END)
        else:
            textPad.insert(CURRENT, str(self.cn[self.order[self.myindex]]) + '\t')
            textPad.see(END)

    def start_cntoen_pronunc_random(self):
        self.myindex = 0
        self.type = 6
        textPad.delete(1.0, END)
        textPad.insert(1.0,'总单词数大约：'+str(self.length)+'\n\n')
        random.shuffle(self.order)
        while(self.myindex<self.length and self.cn[self.order[self.myindex]]==''):    # 不显示无中文翻译的单词
            self.myindex = self.myindex + 1
        if(self.myindex==self.length):
            textPad.insert(CURRENT, '[all done!]' + '\n')
        else:
            textPad.insert(CURRENT,str(self.cn[self.order[self.myindex]])+'\t')

    def cntoen_pronunc_next_random(self,astr):
        if (self.myindex == self.length):
            textPad.insert(CURRENT, '[all done!]' + '\n')
            textPad.see(END)
            return

        if (astr == loophole):
            textPad.insert(CURRENT,'\n')
            textPad.see(END)
        else:
            textPad.insert(CURRENT, astr + '\n')
            textPad.see(END)
        if(astr == loophole or astr==self.en[self.order[self.myindex]]):
            # textPad.insert(CURRENT, '[right!]' + '\n\n')  #without tag
            textPad.insert(CURRENT, "right!", "tag_right")
            textPad.see(END)
            textPad.insert(CURRENT,  "\n\n")
            textPad.see(END)

        else:
            self.times[self.order[self.myindex]] = self.times[self.order[self.myindex]] + 1
            textPad.insert(CURRENT, "wrong!" + '\t' + '累计错误次数:' + str(int(self.times[self.order[self.myindex]])) + ' ',"tag_wrong")
            textPad.see(END)
            textPad.insert(CURRENT,'\t'+str(self.en[self.order[self.myindex]])+'\t'+str(self.pronunc[self.order[self.myindex]])+'\n\n')
            textPad.see(END)
        self.myindex = self.myindex + 1
        while (self.myindex<self.length and self.cn[self.order[self.myindex]]==''):  # 不显示无中文翻译的单词
            self.myindex = self.myindex + 1
        if (self.myindex == self.length):
            textPad.insert(CURRENT, '[all done!]' + '\n')
            textPad.see(END)
        else:
            textPad.insert(CURRENT, str(self.cn[self.order[self.myindex]]) + '\t')
            textPad.see(END)

    def save_log_file(self):
        global myfiles
        # log文件写入与保存
        if (myfiles.log_exist == 0):
            logtable = self.f.add_sheet(u'sheet1', cell_overwrite_ok=True)
            logtable.write(0, 0, 'word')
            logtable.write(0, 1, 'spell_wrong_times')
            logtable.write(0, 4, '(DO NOT CHANGE THIS FILE)')
            logtable.write(1, 4, '(JUST FOR READING)')
            for i in range(0, len(myapp.en)):
                logtable.write(i + 1, 0, myapp.en[i])
            for j in range(0, len(myapp.times)):
                logtable.write(j + 1, 1, myapp.times[j])
            self.f.save(myfiles.file_path + myfiles.log_name)

        else:
            log = xlutils.copy.copy(self.f)
            logtable = log.get_sheet(0)
            logtable.write(0, 0, 'word')
            logtable.write(0, 1, 'spell_wrong_times')
            logtable.write(0, 3, '(DO NOT CHANGE THIS FILE)')
            logtable.write(1, 3, '(JUST FOR READING)')
            for i in range(0, len(myapp.en)):
                logtable.write(i + 1, 0, myapp.en[i])
            for j in range(0, len(myapp.times)):
                logtable.write(j + 1, 1, myapp.times[j])
            log.save(myfiles.file_path + myfiles.log_name)
        self.f = NONE

    def callback(self):
        if askokcancel("退出", "你确定要退出吗?"):
            if(self.f!=NONE):
                self.save_log_file()
            root.destroy()

def author():
    showinfo('作者信息','作者: chenxiaoyu'+'\n'+"邮箱: chenxiaoyu.xintong@gmail.com" )

def use():
    showinfo('使用说明','1.准备词典：词典需为xlsx或xls格式，第一列存英文，第二列存音标（若没有请留空），第三列存中文；第一行存表头，内部会自动切去（若没有请留空）'
                         '\n2.点击文件选项中的读取词典'
                         '\n3.选择功能选项中的对应功能'
                         '\n4.对于显示的单词，请把对应翻译写在下方空框中' )

def about():
    showinfo('当前版本 V4','【更新日志】\n\n'
            '【V4  2016/10/9】\n1.自动记录中->英错误次数'
                                '\n3.加入词典格式检查功能'
                                '\n2.加入强制退出，如遇到bug无法正常退出时可使用此选项\n\n'
            '【V3  2016/10/7】\n1.加入音标显示'
                                '\n2.添加使用说明选项'
                                '\n3.添加随机考察功能'
                                '\n4.修复bug：当文字满屏却未将滚动条滚到最底端，会出现插入错误\n\n'
            '【V2  2016/10/6】\n1.加入音标显示'
                                '\n2.添加使用说明选项'
                                '\n3.添加随机考察功能'
                                '\n4.将默认焦点设为Entry组件\n\n'
            '【V1  2016/10/1】\n1.开启中译英、英译中的判断'
                                '\n2.加入有色标签，用于强化显示判断结果'
                                '\n3.加入一个后台判断漏洞，仅供娱乐(翻译时写入作者名即可见到效果)\n\n'
            '【alpha    2016/9/20】\n1.请先读取xlsx格式的词典，xlsx第一列存英文，第二列存对应中文翻译（第一行可存表头，内部会自动切去）'
                                '\n2.之后在选择功能，可进行中译英，英译中的自查，对结果可保存为txt格式的文本')

def resource():
    showinfo('相关资源','http://www.tkdocs.com/tutorial/index.html and http://effbot.org/tkinterbook/tkinter-index.html')

def shutdown_anyway():
    if askokcancel("强制退出", "你确定要【强制退出】？仅限无法正常退出时使用！"):
        root.destroy()

def read():
    global myfiles
    global myapp
    dict_file=askopenfile()
    myfiles.set(dict_file.name)
    root.title('当前词典：'+str(os.path.basename(dict_file.name)))
    myapp.read_data(dict_file.name)

def saveas():
    f=asksaveasfilename(initialfile='未命名.txt',defaultextension='.txt')
    global filename
    filename = f
    fh=open(f,'w')
    msg=textPad.get(1.0,END)
    fh.write(msg)
    fh.close()
    root.title(os.path.basename(f))

def entocn():
    myapp.start_entocn()

def cntoen():
    myapp.start_cntoen()

def cntoen_pronunc():
    myapp.start_cntoen_pronunc()

def entocn_random():
    myapp.start_entocn_random()

def cntoen_random():
    myapp.start_cntoen_random()

def cntoen_pronunc_random():
    myapp.start_cntoen_pronunc_random()

class dict_and_log(object):
    # 传入词典文件名，带路径，如 C:/Users/chenxy/Desktop/jiajiao/词汇.xlsx
    # filename = askopenfile()
    # print(filename.name)
    def __init__(self):
        self.full = ''
        self.addon = '_Log.xls'
        self.file_path = ''
        self.dict_name = ''
        self.log_name = ''
        self.log_exist = 0

    def set(self,full_name):
        self.full = full_name

        splits_of_full = self.full.split('/')
        for i in range(len(splits_of_full) - 1):
            self.file_path += splits_of_full[i]
            self.file_path += '/'

        splits_of_name = splits_of_full[-1].split('.')
        for i in range(len(splits_of_name) - 1):
            self.dict_name += splits_of_name[i]
            self.dict_name += '.'
        self.dict_name = self.dict_name[:-1]

        self.log_name = self.dict_name + self.addon

        if(os.path.exists(self.file_path+self.log_name)):
            self.log_exist = 1
        else:
            self.log_exist = 0


myfiles = dict_and_log()
myapp = application()

root=Tk()
root.title('单词翻译自查工具')
root.geometry("800x500")


#Create Menu
root.option_add('*tearOff', FALSE)
menubar=Menu(root)
root.config(menu=menubar)
filemenu=Menu(menubar)
menubar.add_cascade(label='文件',menu=filemenu)
filemenu.add_command(label='读取词典',accelerator='Ctrl+R',command=read)
filemenu.add_command(label='另存为',accelerator='Ctrl+S',command=saveas)

editmenu=Menu(menubar)
menubar.add_cascade(label='功能',menu=editmenu)
editmenu.add_command(label='顺序考查：英 -> 中',command=entocn)
editmenu.add_command(label='顺序考查：中 -> 英 (忽略音标)',command=cntoen)
editmenu.add_command(label='顺序考查：中 -> 英 (显示音标)',command=cntoen_pronunc)
editmenu.add_separator()
editmenu.add_command(label='随机考查：英 -> 中',command=entocn_random)
editmenu.add_command(label='随机考查：中 -> 英 (忽略音标)',command=cntoen_random)
editmenu.add_command(label='随机考查：中 -> 英 (显示音标)',command=cntoen_pronunc_random)

aboutmenu=Menu(menubar)
menubar.add_cascade(label='关于',menu=aboutmenu)
aboutmenu.add_command(label='作者信息',command=author)
aboutmenu.add_command(label='使用说明',command=use)
aboutmenu.add_command(label='版本信息',command=about)
aboutmenu.add_command(label='相关资源',command=resource)
aboutmenu.add_command(label='强制退出',command=shutdown_anyway)


#create toolbar
# toolbar=Frame(root,height=25)

# shortbutton=Button(toolbar,text='保存',command=save)
# shortbutton.pack(side=LEFT,padx=5,pady=5)
# toolbar.pack(side=BOTTOM)

# shortbutton=Button(toolbar,text='ss',command=read)
# shortbutton.pack(side=RIGHT,padx=5,pady=5)
# toolbar.pack(side=BOTTOM,expand=NO,fill=Y)

#create statusbar
# status=Label(root,bd=1,relief=SUNKEN,anchor=W,height=1)
# status.pack(side=BOTTOM,fill=X)

#create linenumber & scroll
lnlabel= Label(root,width=2)
lnlabel.pack(side=LEFT,fill=Y)
textPad=Text(root,undo=True)
textPad.pack(expand=YES,fill=BOTH)
# 设置 tag
textPad.tag_config("tag_right", backgroun="DodgerBlue", foreground="yellow")
textPad.tag_config("tag_wrong", backgroun="Red", foreground="yellow")
# 加入loophole
loophole = 'chenxiaoyu'

scroll=Scrollbar(textPad)
textPad.config(yscrollcommand=scroll.set)
scroll.config(command=textPad.yview)
scroll.config(cursor='hand2')
scroll.pack(side=RIGHT,fill=Y)

def rtnkey(event=None):
    global myapp
    if(myapp.type == 1):
        myapp.entocn_next(e.get())
    elif(myapp.type == 2) :
        myapp.cntoen_next(e.get())
    elif(myapp.type == 3):
        myapp.cntoen_pronunc_next(e.get())
    elif (myapp.type == 4):
        myapp.entocn_next_random(e.get())
    elif (myapp.type == 5):
        myapp.cntoen_next_random(e.get())
    elif (myapp.type == 6):
        myapp.cntoen_pronunc_next_random(e.get())
    e.initialize('')

# create Entry module
e = StringVar()
entry = Entry(root, validate='key', textvariable=e, width=20,font=("宋体", 18, "bold"))
entry.pack()
entry.bind('<Return>', rtnkey)
entry.focus_force()

root.protocol("WM_DELETE_WINDOW", myapp.callback)
root.mainloop()



