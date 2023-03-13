# -*- coding : utf-8 -*-
"""
@author: Github-Afool4U
@Description: GPAHelper.py
@Date: 2023-3-13 21:15
"""
import decimal
from time import sleep
from io import StringIO
import ctypes
import threading
from tkinter import Tk
import tkinter.font as tkFont
from tkinter.scrolledtext import ScrolledText
import pyperclip
from csv import DictReader

WIDTH = 550  # 窗口宽
HEIGHT = 300  # 窗口高
header = '学年	学期	课程代码	课程名称	课程性质	课程归属	学分	绩点	成绩	辅修标记	补考成绩	重修成绩	开课学院	备注	重修标记'
tail = '学生条形码'


def resize(text):
    content = text.get('1.0', 'end')
    lines = content.split('\n')
    max_line_len = max(font.measure(line) for line in lines)
    width = max_line_len + text.cget('padx') * 2 + 30
    height = len(lines) * font.metrics('linespace') + font.metrics("ascent") + text.cget('pady') * 2
    root.geometry(f'{width}x{height}')
    root.update()
    text.configure(state='disabled')


def show_result(content, text):
    text.configure(state='normal')
    text.delete('1.0', 'end')
    # 总学分
    credit = decimal.Decimal('0')
    # 总绩点
    gpa = decimal.Decimal('0')
    data_io = StringIO(content)
    sheet = DictReader(data_io, delimiter='\t')
    non_gpa_courses = []
    # 遍历每一门课程，“成绩”为合格的不计入（二级制不算绩点）
    for row in sheet:
        if not row['成绩'].strip() == '' and row['成绩'] != '合格' and not row['绩点'].strip() == '':
            credit += decimal.Decimal(row['学分'])
            gpa += decimal.Decimal(row['学分']) * decimal.Decimal(row['绩点'])
        else:
            non_gpa_courses.append({'课程名称': row['课程名称'], '学分': row['学分'],
                                    '绩点': row['绩点'] if not row['绩点'].strip() == '' else '无数据'})
    text.insert("end", '平均绩点：' + ((gpa / credit).quantize(decimal.Decimal('0.000000'))).__str__())
    text.insert("end", '\n总学分(非考查课且已出GPA)：' + credit.__str__())
    text.insert("end", '\n\n未计算的课程：\n')
    for idx, course in enumerate(non_gpa_courses):
        text.insert("end", f'{idx + 1}. ' + course['课程名称'] + '  ' + course['学分'] + '  ' + course['绩点'] + '\n')
    text.insert("end",
                f'共{len(non_gpa_courses)}门，{sum(float(course["学分"]) for course in non_gpa_courses)}学分\n')
    resize(text)


def check(text):
    recent_value = pyperclip.paste()
    while True:
        tmp_value = pyperclip.paste()
        if tmp_value != recent_value:
            if header in tmp_value and tail in tmp_value:
                raw = tmp_value
                break
            else:
                recent_value = tmp_value
        sleep(0.1)
    raw = raw[raw.find(header):]
    raw = raw[:raw.rfind(tail)]
    raw = raw.strip()
    show_result(raw, text)


def mk_window():  # 创建窗口对象
    root = Tk()
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
    # 调用api获得当前的缩放因子
    ScaleFactor = ctypes.windll.shcore.GetScaleFactorForDevice(0)
    # 设置缩放因子
    root.tk.call('tk', 'scaling', ScaleFactor / 75)
    root.title("GPA计算助手--github.com/Afool4U")
    root.resizable(True, True)
    root.geometry("{}x{}+{}+{}".  # 初始化窗口大小和位置
                  format(WIDTH, HEIGHT, int(root.winfo_screenwidth() * ScaleFactor / 100 - WIDTH) // 2,
                         int(root.winfo_screenheight() * ScaleFactor / 100 - HEIGHT) // 2))
    root.attributes("-topmost", True)
    return root


if __name__ == '__main__':
    # 查询结果用GUI显示
    global font
    root = mk_window()
    font = tkFont.Font(root=root, family='Times', size=17)
    text = ScrolledText(root, width=WIDTH, height=HEIGHT, font=font)
    text.insert("end", '使用方法:\n')  # 插入文本
    text.insert("end", '教务系统-->信息查询-->成绩查询-->学期/学年/历年成绩，鼠标左键单击成绩单中任意一处，然后先按下Ctrl+A，再按下Ctrl+C即可计算。')
    text.configure(state='disabled')
    text.pack()
    threading.Thread(target=check, args=(text,)).start()  # 启动监测线程
    root.mainloop()  # 进入主消息循环
