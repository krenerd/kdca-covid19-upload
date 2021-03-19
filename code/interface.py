import tkinter as tk
from tkinter import filedialog, scrolledtext
from tkinter import *
import os
import write_to_web
import data_refiner
import threading
def refine_data():
    global file_path
    upload['state']='disabled'
    refine['state']='disabled'
    complete['state']='disabled'
    data_refiner.refine_data(file_path,top,text_area)
   
    upload['state']='normal'
def configure_TK():
    top=tk.Tk()
    top.resizable(width=False, height=False)
    top.geometry('350x400')
    top.title('질병관리본부 기입프로그램')
    top.configure(background='#CDCDCD')
    return top

def enter_excel():
    global file_path
    if complete['text']=='데이터 입력':
        upload['state']='disabled'
        refine['state']='disabled'
        write_to_web.open_webbrowser()
        complete['text']='공인인증서 완료'
    else:
        complete['state']='disabled'
        write_to_web.enter_excel(file_path,top,text_area,entry1.get())
        complete['text']='데이터 입력'
        upload['state']='normal'
       
def valid_file_path(file_path):
    file_type=file_path.split('.')[-1]
    if file_type=='xlsx':
        return True
    else:
        return False
def upload_image():
    try:
        global file_path
        file_path=filedialog.askopenfilename()
        if valid_file_path(file_path):
            upload['state']='disabled'
            refine['state']='normal'
            complete['state']='normal'
    except:
        pass

top=configure_TK()
label=Label(top,background='#CDCDCD', font=('arial',15,'bold'))

upload=Button(top,text="엑셀 파일 선택",command=upload_image,
  padx=10,pady=5)
upload.configure(background='#364156', foreground='white',
    font=('arial',10,'bold'))

complete=Button(top,text="데이터 입력",command=enter_excel,
      padx=10,pady=5)
complete.configure(background='#364156', foreground='white',
      font=('arial',10,'bold'))

refine=Button(top,text="데이터 수정",command=refine_data,
      padx=10,pady=5)
refine.configure(background='#364156', foreground='white',
      font=('arial',10,'bold'))

entry1 = tk.Entry(top)

text_area = scrolledtext.ScrolledText(top,  wrap = tk.WORD, width = 50, height = 11,font = ("Times New Roman",10))
 
label.pack(side=BOTTOM,expand=True)
text_area.pack(side=BOTTOM)
entry1.pack(side=BOTTOM,pady=8)
complete.pack(side=BOTTOM)
refine.pack(side=BOTTOM,pady=8)
upload.pack(side=BOTTOM)

heading = Label(top, text="질병관리본부 기입프로그램",pady=20, font=('arial',20,'bold'))
entry1.insert(END, '31700543')

heading.configure(background='#CDCDCD',foreground='#364156')
heading.pack()

refine['state']='disabled'
complete['state']='disabled'
text_area['state']='disabled'
top.mainloop()
