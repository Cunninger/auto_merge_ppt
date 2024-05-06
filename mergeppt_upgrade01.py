import os

import win32com.client as win32
from tkinter import Tk, Label, Button, Entry, StringVar, Toplevel, Text


def join_ppt(path: str, save_path: str):
    files = os.listdir(path)
    files.sort(key=lambda x: os.path.getmtime(os.path.join(path, x)), reverse=True)
    Application = win32.gencache.EnsureDispatch("PowerPoint.Application")

    Application.Visible = 1
    new_ppt = Application.Presentations.Add()

    for file in files:
        abs_path = os.path.join(path, file)
        exit_ppt = Application.Presentations.Open(abs_path)
        print('正在操作的文件：', abs_path)
        page_num = exit_ppt.Slides.Count
        exit_ppt.Close()
        new_ppt.Slides.InsertFromFile(abs_path, new_ppt.Slides.Count, 1, page_num)
    new_ppt.SaveAs(save_path)
    Application.Quit()


class PPTJoinerApp:
    def __init__(self, master):

        self.master = master
        master.title("PPT Joiner(合并前请先打开PowerPoint!)")
        # 定义框的大小
        master.geometry('400x200')
#
        self.source_label = Label(master, text="源文件位置")
        self.source_label.pack()

        self.source_var = StringVar()
        self.source_entry = Entry(master, textvariable=self.source_var,width=50)
        self.source_entry.pack()
#

        self.save_label = Label(master, text="保存位置")
        self.save_label.pack()

        self.save_var = StringVar()
        self.save_entry = Entry(master, textvariable=self.save_var,width=50)
        self.save_entry.pack()


        self.join_button = Button(master, text="合并", command=self.join_ppt, width=25)

        self.join_button.pack()

        self.doc_button = Button(master, text="说明文档", command=self.open_documentation, width=10)
        self.doc_button.pack()

    def join_ppt(self):
        source_path = self.source_var.get().replace('"', '').replace('\\', '/')
        save_path = self.save_var.get().replace('"', '').replace('\\', '/')

        join_ppt(source_path, save_path)

    def open_documentation(self):
        doc_window = Toplevel(self.master)
        doc_window.title("说明文档")

        doc_text = Text(doc_window)
        doc_text.insert('1.0',
                        '源文件位置说明：\n 右键文件夹，找到并点击复制文件地址，粘贴到源文件位置\n\n' '保存位置说明：\n 合并好的ppt文件默认放在桌面,命名为merge.pptx\n\n' '点击合并按钮，合并成功后会在保存位置生成一个新的ppt文件\n\n' '注意：\n ppt文件名不要有空格，否则会报错')
        doc_text.pack()


root = Tk()
app = PPTJoinerApp(root)
root.mainloop()
