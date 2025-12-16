from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import main as Main

root = Tk()
root.geometry("700x475")
root.title('CA & Rollout Bot')
root.iconbitmap('Images\\huawei_icon.ico')
root.iconify()
root.configure(bg = 'white')
logo = PhotoImage(file = "Images\\auditor3.png")
logo_lbl = Label(root, image = logo)
logo_lbl.pack()

def CE():
    global ce
    ce = filedialog.askopenfilename(title = 'Select "CA" File')
    CE_lbl = Label(root, text=ce, bg = 'azure2', font = ("Times New Roman", 10, 'bold'))
    CE_lbl.place(x=120, y=165)

def CODE():
    global code
    code = filedialog.askopenfilename(title = 'Select "RollOut" file')
    CODE_lbl = Label(root, text=code, bg = 'azure2', font = ("Times New Roman", 10, 'bold'))
    CODE_lbl.place(x=120, y=215)



def output():
    global out_path
    out_path = filedialog.askdirectory(title = 'Select Output directory')
    out_path_lable = Label(root, text=out_path, bg = 'azure2', font = ("Times New Roman", 10, 'bold'))
    out_path_lable.place(x=120, y=275)

def close():
    gh = messagebox.askquestion('Warning', 'Are you sure you want to quit?')
    if gh == 'yes':
        root.quit()


btn_freq = Button(root, text='CA File', bg='lavender', font=("Times New Roman", 10, "bold"), width=10, command=lambda: CE())
btn_freq.place(x=30, y=160)

btn_map = Button(root, text='RO File', bg='lavender', font=("Times New Roman", 10, "bold"), width=10, command=lambda: CODE())
btn_map.place(x=30, y=210)

btn_out = Button(root, text='Output', bg='lavender', font=("Times New Roman", 10, "bold"), width=10, command=lambda: output())
btn_out.place(x=30, y=270)


try:
    btn_start = Button(root, text='Start', bg='lavender', font=("Times New Roman", 12, "bold"), width=6, command=lambda:Main.OD_IBS_SLA_STATUS(ce, code, out_path))
except:
    btn_start = Button(root, text='Start', bg='lavender', font=("Times New Roman", 12, "bold"), width=6, command=lambda:Main.OD_IBS_SLA_STATUS('ce', code, out_path))

# btn_start = Button(root, text='Start', bg='lavender', font=("Times New Roman", 12, "bold"), width=6, command=lambda:Main.OD_IBS_SLA_STATUS(input_CE, input_Code, input_TCP, input_DL_UL, required_file, Output_path))
btn_start.place(x=220, y=360)

btn_quit = Button(root, text='Quit', bg='lavender', font=("Times New Roman", 12, "bold"), width=6, command=lambda: close())
btn_quit.place(x=310, y=360)

lbl_signature1 = Label(root, text='               For Support :  Danish Ali (dwx854280)\n             Contact :  00971508552942', bg='white', font=("Times New Roman", 8, 'bold'))
lbl_signature1.place(x=150, y=430)

root.mainloop()
