import tkinter as tk
window = tk.Tk()
window.resizable(0,0)
window.title('黃金比例計算器')
window.geometry('300x200+300+200')
label1=tk.Label(window, text='請輸入數字')
label2=tk.Label(window, text='x1.618')
label3=tk.Label(window, text='x0.618')
input=tk.Entry(window,width=10)
result1=tk.Entry(window,width=10)
result1.insert(0,'0')
result2=tk.Entry(window,width=10)
result2.insert(0,'0')
r=(1+(5)**0.5)/2

def click(event=None):
    k=input.get()
    k1=float(k)*r
    k2=float(k)*(1/r)
    result1.delete(0,"end")
    result1.insert(0, '%.3f'%k1)
    result2.delete(0,"end")
    result2.insert(0, '%.3f'%k2)
    
button=tk.Button(window,text='計算',command=click)
window.bind('<Return>',click)
label1.pack(side='top')
input.pack(side='top')
label2.pack(side='top')
result1.pack(side='top')
label3.pack(side='top')
result2.pack(side='top')
button.pack(side='top')
window.mainloop()