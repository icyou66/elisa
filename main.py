import datetime
import glob
import os
import time
import tkinter
import traceback
from threading import Thread
from tkinter import *
from tkinter import messagebox

import openpyxl

from elisa import Elisa


class Tkinter:
    def __init__(self):
        self.backup = None
        self.log_place = None
        self.log_data = None
        self.root = Tk()
        self.color = ['red', 'green', 'blue', 'saddlebrown', 'yellow', 'cyan', 'purple']
        self.root.title("ELISA数据处理工具")
        window_height = self.root.winfo_screenheight()
        window_width = self.root.winfo_screenwidth()
        width, height = list([300, 500])
        self.root.geometry(
            "%dx%d+%d+%d"
            % (width, height, (window_width - width) / 2, (window_height - height) / 2)
        )
        self.root.attributes("-alpha", 1)

        self.create_place()  # 创建窗口
        self.root.mainloop()  # 使窗口等待

    def create_place(self):
        self.info_place_func()
        self.log_place_func()

    def info_place_func(self):
        place = Frame(self.root)
        place.pack(side="top", fill=BOTH)
        info_place = LabelFrame(place)
        info_place.pack(ipadx=5, ipady=5, fill="x", padx=10)

        text = "使用教程：将Excel数据表格放在程序的同一目录下，然后点击开始处理即可"
        Label(info_place, text=text, font="微软雅黑 9", fg="blue", wraplength=260).pack()
        Button(info_place, text="开始处理", command=lambda: self.run_thread(self.run), relief="groove",
               background="#C0E3FF").pack(padx=10, pady=5, fill=tkinter.X)

    def log_place_func(self):
        # 日志空间
        self.log_place = Frame(self.root)
        self.log_place.pack(expand=tkinter.YES, fill=tkinter.BOTH, side="bottom", padx=5, pady=5, ipadx=5, ipady=5)
        self.log_data = Text(self.log_place, font="微软雅黑 9", height=16, fg="blue", bg="#F4F3FF")
        self.log_data.pack(expand=tkinter.YES, fill=tkinter.BOTH, side="bottom", padx=5, pady=5, ipadx=5, ipady=5)
        for item in self.color:
            self.log_data.tag_config(item, foreground=item)
        self.pprint("Tip：操作后这里显示日志\n")
        self.pprint("当前目录下Excel文件如下：")
        xlsx_files = glob.glob(os.path.join(os.getcwd(), '*.xlsx'))
        if not xlsx_files:
            self.pprint("无Excel文件！请检查！", color="red")
        for i, file in enumerate(xlsx_files):
            self.pprint(f"{i + 1}，{file.split("\\")[-1]}", color="purple")

        roll_logy = tkinter.Scrollbar(self.log_data)
        roll_logy.pack(side="right", fill="y")
        self.log_data.config(yscrollcommand=roll_logy.set)
        roll_logy.config(command=self.log_data.yview)

    @staticmethod
    def run_thread(func):
        Thread(target=func).start()

    def run(self):
        self.log_data.delete("1.0", END)
        self.pprint("======程序运行======")
        xlsx_files = glob.glob(os.path.join(os.getcwd(), '*.xlsx'))
        xlsx_files = [file for file in xlsx_files if "~$" not in file]
        if not xlsx_files:
            messagebox.showwarning("提示", "目录下暂无Excel文件！请检查！")
            return

        for path in xlsx_files:
            # 获取Excel工作簿对象
            excel_handle = openpyxl.load_workbook(path)
            if excel_handle.worksheets[0]['A15'].value == "完成后弹出板":
                self.pprint(f"ELISA数据文件：{path.split('\\')[-1]}\n", color="saddlebrown")
                time.sleep(1)
                Elisa(excel_handle, self.pprint).start()
            else:
                self.pprint(f"未知文件：{path.split('\\')[-1]}\n", color="red")
                continue

            try:
                excel_handle.save(path)
            except PermissionError:
                self.pprint("该Excel文件在你电脑上已被打开，请关闭后重新运行程序！", color="red")
                return False
            except:
                self.pprint(traceback.format_exc(), color="red")
                self.pprint("保存文件时发生异常错误，请截图联系管理员！", color="red")
                return False

        self.pprint(f"[{self.fetch_time()}]所有文件处理完毕，程序运行结束！", color="green")

    @staticmethod
    def fetch_time():
        current_time = datetime.datetime.now()
        formatted_time = current_time.strftime("%H:%M:%S")
        return formatted_time

    def pprint(self, msg, color=None, delete=False):
        if not color:
            color = "blue"
        if delete:
            start = self.log_data.index("end-2l")
            end = self.log_data.index("end-1c")
            self.log_data.delete(start, end)
        self.log_data.insert(END, msg + "\n", color)
        self.log_data.update()
        self.log_data.see(END)

    def end(self, msg, color=None, delete=False):
        if not color:
            color = "blue"
        if delete:
            start = self.log_data.index("end-2l")
            end = self.log_data.index("end-1c")
            self.log_data.delete(start, end)
        self.log_data.insert(END, msg + "\n", color)
        self.log_data.update()
        self.log_data.see(END)
        raise SystemExit


if __name__ == "__main__":
    Tkinter()
