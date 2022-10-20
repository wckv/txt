import tkinter as tk
import datetime

class App:
    def __init__(self, root):
        #setting title
        root.title("申報轉換")
        #setting window size
        width=630
        height=393
        screenwidth = root.winfo_screenwidth()
        screenheight = root.winfo_screenheight()
        alignstr = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
        root.geometry(alignstr)
        root.resizable(width=False, height=False)

        global GLineEdit_708,GLineEdit_695,GLineEdit_316

        GLineEdit_708=tk.Text(root)
        GLineEdit_708["font"]=('微軟正黑體',10)
        GLineEdit_708["borderwidth"] = "1px"
        GLineEdit_708.place(x=10,y=10,width=370,height=370)

        GLineEdit_695=tk.Entry(root)
        GLineEdit_695["font"]=('Arial',10)
        GLineEdit_695["borderwidth"] = "1px"
        GLineEdit_695["justify"] = "center"
        GLineEdit_695.insert(0,"W005595")
        GLineEdit_695.place(x=558,y=180,width=62,height=38)

        GLineEdit_316=tk.Entry(root)
        GLineEdit_316["font"]=('Arial',10)
        GLineEdit_316["borderwidth"] = "1px"
        GLineEdit_316["justify"] = "center"
        GLineEdit_316.insert(0,"GC4219")
        GLineEdit_316.place(x=558,y=230,width=62,height=38)

        GButton_790=tk.Button(root)
        GButton_790["font"]=('微軟正黑體',14)
        GButton_790["justify"] = "center"
        GButton_790["text"] = "彙總"
        GButton_790.place(x=390,y=280,width=161,height=100)
        GButton_790["command"] = self.GButton_790_command

        GLabel_217=tk.Label(root)
        GLabel_217["font"]=('微軟正黑體',12)
        GLabel_217["anchor"]='e'
        GLabel_217["text"] = "Molnupiravir批號"
        GLabel_217.place(x=390,y=185,width=150,height=30)

        GLabel_617=tk.Label(root)
        GLabel_617["font"]=('微軟正黑體',12)
        GLabel_617["anchor"]='e'
        GLabel_617["text"]="Paxlovid批號"
        GLabel_617.place(x=390,y=235,width=150,height=30)

        GMessage_355=tk.Label(root)
        GMessage_355["font"]=('微軟正黑體',12)
        GMessage_355["anchor"]='nw'
        GMessage_355["justify"]="left"
        GMessage_355["text"]='請從Excel複製病人資料\n從藥名框到庫台\n範例：'
        GMessage_355.place(x=390,y=10,width=250,height=100)

        GMessage_35=tk.Label(root)
        GMessage_35["font"]=('微軟正黑體',10)
        GMessage_35["anchor"]='nw'
        GMessage_35["justify"]="left"
        GMessage_35["text"]='Molnupiravir 包龍星 A123456789 NA\nMolnupiravir \
戚秦氏 B234567890 NB\nPaxlovid 豹子頭 C123454321 NP\nPaxlovid 常威 D1357997531 NE'
        GMessage_35.place(x=393,y=75,width=250,height=100)



        GButton_216=tk.Button(root,font=('微軟正黑體',14))
        GButton_216["justify"] = "center"
        GButton_216["text"] = "清除"
        GButton_216.place(x=560,y=280,width=62,height=100)
        GButton_216["command"] = self.GButton_216_command

    def GButton_790_command(self):
        lst=GLineEdit_708.get(1.0,'end')
        lst=lst.split()
        M=[]
        P=[]
        for i in range(len(lst)):
            if lst[i]=='Molnupiravir':
                M+=['【'+' '.join(lst[i+1:i+4])+'】']
            elif lst[i]=='Paxlovid':
                P+=['【'+' '.join(lst[i+1:i+4])+'】']
        nM=lst.count('Molnupiravir')
        nP=lst.count('Paxlovid')
        today=datetime.date.today()
        yesterday=today+datetime.timedelta(-1)
        today=today.strftime('%Y/%m/%d')
        yesterday=yesterday.strftime('%m/%d')
        date=today+'申報'+yesterday+'處方'
        
        Res=date+'【Molnupiravir批號'+GLineEdit_695.get()+'=>'+str(nM)+'PC】'
        Res+=''.join(M)+'\n'*5
        Res+=date+'【Paxlovid批號'+GLineEdit_316.get()+'=>'+str(nP)+'PC】'
        Res+=''.join(P)
        GLineEdit_708.delete(1.0,'end')
        GLineEdit_708.insert(1.0,Res)

    def GButton_216_command(self):
        GLineEdit_708.delete(1.0,'end')

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
