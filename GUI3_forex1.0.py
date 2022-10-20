from re import M
import tkinter as tk
import tkinter.font as tkFont
from tkinter import StringVar

class App:
    def __init__(self, root):
        #setting title
        root.title("客家一號")
        #setting window size
        width=480
        height=640
        screenwidth = root.winfo_screenwidth()
        screenheight = root.winfo_screenheight()
        alignstr = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
        root.geometry(alignstr)
        root.resizable(width=False, height=False)

        global GLineEdit_421,GLineEdit_783,GLineEdit_334,k

        GLineEdit_421=tk.Entry(root)
        GLineEdit_421["borderwidth"] = "1px"
        ft = tkFont.Font(family='Times',size=16)
        GLineEdit_421["font"] = ft
        GLineEdit_421["fg"] = "#333333"
        GLineEdit_421["justify"] = "center"
        #GLineEdit_421["text"] = '0.2222'
        GLineEdit_421.place(x=60,y=10,width=164,height=30)

        GLabel_781=tk.Label(root)
        ft = tkFont.Font(family='微軟正黑體',size=14)
        GLabel_781["font"] = ft
        GLabel_781["fg"] = "#333333"
        GLabel_781["justify"] = "center"
        GLabel_781["text"] = "匯率"
        GLabel_781.place(x=0,y=15,width=60,height=25)

        GLineEdit_783=tk.Entry(root)
        GLineEdit_783["borderwidth"] = "1px"
        ft = tkFont.Font(family='Times',size=16)
        GLineEdit_783["font"] = ft
        GLineEdit_783["fg"] = "#333333"
        GLineEdit_783["justify"] = "center"
        #GLineEdit_783["text"] = '100'
        GLineEdit_783.place(x=360,y=10,width=100,height=30)

        GLabel_912=tk.Label(root)
        ft = tkFont.Font(family='微軟正黑體',size=14)
        GLabel_912["font"] = ft
        GLabel_912["fg"] = "#333333"
        GLabel_912["justify"] = "center"
        GLabel_912["text"] = "目標價格(±10)"
        GLabel_912.place(x=230,y=10,width=135,height=39)

        k = StringVar()
        GLineEdit_334=tk.Label(root,textvariable=k)
        GLineEdit_334["borderwidth"] = "1px"
        ft = tkFont.Font(family='Times',size=16)
        GLineEdit_334["font"] = ft
        GLineEdit_334['anchor']='nw'
        GLineEdit_334["fg"] = "#333333"
        GLineEdit_334["justify"] = "left"
        GLineEdit_334["text"] = 100
        GLineEdit_334.place(x=10,y=50,width=459,height=578)


        GLineEdit_421.bind('<KeyRelease>',App.run)
        GLineEdit_783.bind('<KeyRelease>',App.run)

    def run(event):
        r=float(GLineEdit_421.get())
        m=int(GLineEdit_783.get())
        d=r*m
        k.set(d)




if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
