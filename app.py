#!/usr/bin/env python
# -*- coding: utf-8 -*-

import tkinter as tk
from tkinter import *
from tkinter import filedialog
import hashlib,pandas as pd
import time,os

LOG_LINE_NUM = 0

class MY_GUI():
    def __init__(self,init_window_name):
        self.init_window_name = init_window_name
    #设置窗口
    def set_init_window(self):
        self.file_path = ""
        self.init_window_name.title("excel处理工具_v1.0")           #窗口名
        #self.init_window_name.geometry('320x160+10+10')                         #290 160为窗口大小，+10 +10 定义窗口弹出时的默认展示位置
        self.init_window_name.geometry('1068x681+10+10')
        #self.init_window_name["bg"] = "pink"                                    #窗口背景色，其他背景色见：blog.csdn.net/chl0000/article/details/7657887
        #self.init_window_name.attributes("-alpha",0.9)                          #虚化，值越小虚化程度越高
        # 消息框
        self.file_name = StringVar()
        self.file_name.set("")
        self.message_text_label = Label(self.init_window_name, textvariable=self.file_name)
        self.message_text_label.grid(row=1,column=1)
        # 输入框
        self.sheet_name_label = Entry(self.init_window_name, text="Sheet名称")
        self.sheet_name_label.grid(row=2, column=1)
        #标签
        self.sheet_label = Label(self.init_window_name, text="请输入sheet名称: ")
        self.sheet_label.grid(row=2, column=0)
        self.result_data_label = Label(self.init_window_name, text="输出结果")
        self.result_data_label.grid(row=0, column=12)
        self.log_label = Label(self.init_window_name, text="日志显示：")
        self.log_label.grid(row=3, column=0)
        #文本框
        self.result_data_Text = Text(self.init_window_name, width=70, height=49)  #处理结果展示
        self.result_data_Text.grid(row=1, column=12, rowspan=15, columnspan=10)
        self.log_data_Text = Text(self.init_window_name, width=66, height=43)  # 日志框
        self.log_data_Text.grid(row=4, column=0, columnspan=10)
        #按钮
        self.select_file_button = Button(self.init_window_name, text="文件上传", bg="lightblue", width=10,command=self.select_file)  # 调用内部方法  加()为直接调用
        self.select_file_button.grid(row=1, column=0)
        self.submit_button = Button(self.init_window_name, text="提交", bg="lightblue", width=10,command=self.submit)  # 调用内部方法  加()为直接调用
        self.submit_button.grid(row=2, column=2)

    #获取当前时间
    def get_current_time(self):
        current_time = time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time()))
        return current_time

    # 文件选择
    def select_file(self):
        select_path = tk.filedialog.askopenfilename()
        self.file_path = select_path
        fileName = os.path.basename(select_path)
        self.file_name.set(fileName)
        print(self.file_path)
        print(self.file_name)

    # 提交
    def submit(self):
        #文件路径
        file_path = self.file_path
        if file_path == "" or file_path == None:
            self.write_log_to_Text("请上传文件")
            return
        # sheet名称
        sheet_name = self.sheet_name_label.get()
        if sheet_name == "" or sheet_name == None:
            self.write_log_to_Text("请确定sheet名称")
            return
        # print("sheet名称：{}".format(sheet_name))
        try:
            content_list = self.file_upload(file_path,sheet_name)
            self.result_data_Text.delete(1.0,END)
            for content in content_list:
                print(content)
                # self.write_log_to_Text(content)
                self.result_data_Text.insert(1.0,content+"\n")
        except Exception as e:
            self.write_log_to_Text(e)

    # 文件上传
    def file_upload(self,file_path,sheet_name):
        df = pd.read_excel(file_path,sheet_name) ##默认读取sheet = 0的
        columns = df.columns.values.tolist() ### 获取excel 表头 ，第一行
        contents = []
        # 把所有单元格内容加入数组
        for idx, row in df.iterrows(): ### 迭代数据 以键值对的形式 获取 每行的数据
            for column in columns:
                content = row[column]
                content = str(content).encode('UTF-8')
                # 不是空不是数字不是英文
                if content != "" and not content.isalpha() and not content.isdigit():
                    contents.append(content.decode("UTF-8"))
        result = []
        # 把数组遍历并获取每个元素出现的次数，把大于等于3次的保存起来
        for i in contents:
            num = contents.count(i)
            if num >= 3:
                result.append("{} : {}".format(i,num))
        # 数组去重
        return set(result)

    #日志动态打印
    def write_log_to_Text(self,logmsg):
        global LOG_LINE_NUM
        current_time = self.get_current_time()
        logmsg_in = str(current_time) +" " + str(logmsg) + "\n"      #换行
        if LOG_LINE_NUM <= 7:
            self.log_data_Text.insert(END, logmsg_in)
            LOG_LINE_NUM = LOG_LINE_NUM + 1
        else:
            self.log_data_Text.delete(1.0,2.0)
            self.log_data_Text.insert(END, logmsg_in)


def gui_start():
    init_window = Tk()              #实例化出一个父窗口
    ZMJ_PORTAL = MY_GUI(init_window)
    # 设置根窗口默认属性
    ZMJ_PORTAL.set_init_window()
    init_window.mainloop()          #父窗口进入事件循环，可以理解为保持窗口运行，否则界面不展示


gui_start()