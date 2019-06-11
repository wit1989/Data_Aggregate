#!/usr/bin/env python3
# -*- coding:utf-8 -*-

#!/usr/bin/env python3
# -*- coding:utf-8 -*-


import tkinter.filedialog as filedialog
from tkinter import *
import tkinter.messagebox
import os, xlrd, xlwt



def filedir():
    global path
    print('按键已被点击')
    v.set('')  # 清空文本框里内容
    var.set((('')))
    path = filedialog.askdirectory()
    # print(dir(filedialog))
    if path:
        v.set(path)
    getdir(path)


def getdir(p):
    global fp
    # 把目录中遍历出来的文件目录显示到列表框中
    fp = os.listdir(p)
    print(fp)
    var.set(fp)


def agg_excel():
    global input_sheethead, path
    if isinstance(int(input_sheethead.get()), int):
        sheethead = int(input_sheethead.get())
    else:
        sheethead = 0
    print(sheethead, type(sheethead))
    # path = '/Users/cbowen/downloads/分院上报材料/all'
    os.chdir(path)

    all_wb = xlwt.Workbook(encoding="utf-8", style_compression=0)
    all_sheet = all_wb.add_sheet("各学院数据汇总")

    max_row = 0
    for file in data_files():
        sub_wb = xlrd.open_workbook(file)
        sub_sheet = sub_wb.sheet_by_index(0)

        for row in range(sub_sheet.nrows - sheethead):
            for column in range(sub_sheet.ncols):
                all_sheet.write(max_row + row, column, sub_sheet.cell_value(row + sheethead, column))
            row += 1
        max_row += sub_sheet.nrows - sheethead

    last_name = file_name('alldata.xls')
    all_wb.save(last_name)

    tkinter.messagebox.showinfo('提示', '合并完成,请查看《' + str(last_name) + '》')

def input_num(en=None):
    global input_sheethead
    input_sheethead.delete('0', 'end')


def file_name(name):
    """保存时重名检测"""
    if (name) not in os.listdir('.'):  # 如果本地目录下有重名文件，则在文件名后面加"(n)"，n顺次加1
        even_name = name
    else:
        file_num = 1
        while True:
            dot = name.rfind('.')  # 查询文件名中最后一个'.'的索引
            head = name[:dot]
            tail = name[dot:]

            new_name = head + '(' + str(file_num) + ')' + tail
            if new_name not in os.listdir('.'):
                even_name = new_name
                break
            else:
                file_num += 1
    return even_name


def data_files():
    """提取当前目录下需要合并的文件名，放入列表中"""
    dadafiles = []
    for file in os.listdir('.'):
        if file[-4:] == '.xls' and file[:7] != 'alldata':
            dadafiles.append(file)
    return dadafiles



root = Tk()
root.title('合并Excel')

frame = Frame(root)
frame.pack(fill=X, side=TOP)
# 加入一个文本框显示目录地址
v = StringVar()  # 绑定文本框的变量
ent = Entry(frame, width=50, textvariable=v).pack(fill=X, side=LEFT)
# 加入一个按键，点击后弹出文件目录选择对话框
button = Button(frame, text='选择文件夹', command=filedir).pack(fill=X, side=LEFT)
# 加入一个列表框，显示目录中的文件列表
listframe = Frame(root)
listframe.pack(fill=X, side=TOP)
var = StringVar()  # 绑定listbox的列表值
var.set((''))
listbox = Listbox(listframe, width=60, listvariable=var).pack()

aggfram = Frame(root)
aggfram.pack(fill=X, side=BOTTOM)


input_sheethead = Entry(aggfram, width=50)
input_sheethead.pack(fill=X, side=LEFT)
input_sheethead.insert(0, '此处输入表头行数（双击清空）')
input_sheethead.bind('<Double-1>', input_num)

aggbutton = Button(aggfram, text='点我开始合并', command=agg_excel).pack(fill=X, side=LEFT)


tkinter.messagebox.showinfo('提示', '1、将需要合并的文件放在同一文件夹下；\n2、文件格式为.xls；\n3、文件表头行数要相等。')
root.mainloop()



