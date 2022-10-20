import tkinter as tk

class App:
    def __init__(self, root):
        #setting title
        root.title("廖金順1.0")
        #setting window size
        width=480
        height=640
        screenwidth = root.winfo_screenwidth()
        screenheight = root.winfo_screenheight()
        alignstr = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
        root.geometry(alignstr)
        root.resizable(width=False, height=False)

        GLabel_781=tk.Label(root)
        GLabel_781["font"] = ('微軟正黑體',14)
        GLabel_781["fg"] = "#333333"
        GLabel_781["justify"] = "center"
        GLabel_781["text"] = "匯率"
        GLabel_781.place(x=0,y=15,width=60,height=25)

        GLabel_912=tk.Label(root)
        GLabel_912["font"] = ('微軟正黑體',14)
        GLabel_912["justify"] = "center"
        GLabel_912["text"] = "目標台幣(±20)"
        GLabel_912.place(x=230,y=10,width=135,height=39)

        global rate,amount,show,rLeft,rRight
        #左輸入框rate
        rate=tk.StringVar()
        GLineEdit_421=tk.Entry(root)
        GLineEdit_421['textvariable']=rate
        GLineEdit_421["borderwidth"] = "1px"
        GLineEdit_421["font"] = ('Arial',16)
        GLineEdit_421["justify"] = "center"
        GLineEdit_421.place(x=60,y=10,width=164,height=30)

        #右輸入框amount
        amount=tk.StringVar()
        GLineEdit_783=tk.Entry(root)
        GLineEdit_783['textvariable']=amount
        GLineEdit_783["borderwidth"] = "1px"
        GLineEdit_783["font"] = ('Arial',16)
        GLineEdit_783["justify"] = "center"
        GLineEdit_783.place(x=360,y=10,width=100,height=30)

        #顯示區域show
        show=tk.StringVar()
        GLineEdit_334=tk.Label(root)
        GLineEdit_334["borderwidth"] = "1px"
        GLineEdit_334["font"] = ('微軟正黑體',16)
        GLineEdit_334['anchor']='nw'
        GLineEdit_334["justify"] = "left"
        GLineEdit_334["textvariable"] = show
        GLineEdit_334.place(x=10,y=50,width=460,height=60)

        #左結果rLeft
        rLeft=tk.StringVar()
        GLineEdit_335=tk.Label(root)
        GLineEdit_335["borderwidth"] = "1px"
        GLineEdit_335["font"] = ('微軟正黑體',16)
        GLineEdit_335['anchor']='nw'
        GLineEdit_335["justify"] = "left"
        GLineEdit_335["textvariable"] = rLeft
        GLineEdit_335.place(x=15,y=100,width=220,height=550)


        #右結果rRight
        rRight=tk.StringVar()
        GLineEdit_336=tk.Label(root)
        GLineEdit_336["borderwidth"] = "1px"
        GLineEdit_336["font"] = ('微軟正黑體',16)
        GLineEdit_336['anchor']='nw'
        GLineEdit_336["justify"] = "left"
        GLineEdit_336["textvariable"] = rRight
        GLineEdit_336.place(x=245,y=100,width=220,height=550)

        rate.trace('w',App.calc)
        amount.trace('w',App.calc)

    def calc(*args):
        r=rate.get()
        a=amount.get()
        try:
            r=abs(float(r))
            a=abs(int(a))
            Res='計算：匯率設定'+str(r)+'，計算'+str(a)+'附近的客家值\n'
            Res+='--------------------------------------------------\n'
            fm=(a-20)/r
            if fm<0:
                fm=0
            fM=(a+20)/r
            t=0
            if r>1:
                r/=100
                t=1
            R0=[]
            if t==0:
                for f in range(int(fM-fm)):
                    if 0.39<((f+fm)*r-int((f+fm)*r))<0.5:
                        R0+=[str(format((f+fm),'0.0f'))+' 換 '+str(format((f+fm)*r,'0.3f'))+'元']
            elif t==1:
                k=int((fM-fm)*100)
                for i in range(k):
                    s=(fm+i/100)*r*100
                    if 0.39<s-int(s)<0.5:
                        R0+=[str(format((fm+i/100),'0.2f'))+' 換 '+str(format(s,'0.2f'))+'元']
            rL=R0[:19]
            rL='\n'.join(rL)
            rR=R0[19:38]
            if len(R0)>38:
                rR+=['...下略']
            rR='\n'.join(rR)
            show.set(Res)
            rLeft.set(rL)
            rRight.set(rR)
        except:
            show.set('請在上面輸入匯率與目標價格')
            rLeft.set('')
            rRight.set('')

            






if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
