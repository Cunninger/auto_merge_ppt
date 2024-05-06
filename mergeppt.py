import os
import re
import win32com.client as win32

def join_ppt(path:str):
    """
    :param path: ppt所在文件路径
    :return: None
    """
    files = os.listdir(path)
    # 按照修改时间排序
    files.sort(key=lambda x: os.path.getmtime(os.path.join(path, x)), reverse=True)
    Application = win32.gencache.EnsureDispatch("PowerPoint.Application")

    Application.Visible = 1
    new_ppt = Application.Presentations.Add()
    #执行合并操作
    for file in files:
        abs_path = os.path.join(path, file)
        exit_ppt = Application.Presentations.Open(abs_path)
        print('正在操作的文件：', abs_path)
        page_num = exit_ppt.Slides.Count
        exit_ppt.Close()
        new_ppt.Slides.InsertFromFile(abs_path, new_ppt.Slides.Count, 1, page_num)
    new_ppt.SaveAs("C:/Users/86180/Desktop/merge.pptx") # 括号内为保存位置：如C:\Users\Administrator\Documents\下
    Application.Quit()

join_ppt(r"C:/Users/86180/Desktop/ppt_union")