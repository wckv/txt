from datetime import date,timedelta
from openpyxl import load_workbook

try:
    with open('公費克流感使用清單.txt','r',encoding='big5',errors='ignore') as input:
        input=input.read()[101:]
        n=int(input.count('#')/19)
        input=input.split('#')
        name=input[1:19*n:19]
        ID=input[4:19*n:19]
        drug=input[8:19*n:19]
        lst=list(zip(name,ID,drug))
    lstM=[]
    for i in range(len(lst)):
        if 'P4A151M' in lst[i]:
            lstM+=[lst[i]]
    nameM=[]
    IDM=[]
    for i in lstM:
        nameM+=[i[0]]
        IDM+=[i[1]]

    lstP=[]
    for i in range(len(lst)):
        if 'P4A152M' in lst[i]:
            lstP+=[lst[i]]
    nameP=[]
    IDP=[]
    for i in lstP:
        nameP+=[i[0]]
        IDP+=[i[1]]

    today=date.strftime(date.today(),'%m%d')
    filename='NB_COVID_'+str(today)+'.xlsx'
    wb=load_workbook(filename=filename)
    ws0=wb.worksheets[0]
    ws1=wb.worksheets[1]
    yesterday=date.today()+timedelta(days=-1)
    t0='NB'+date.strftime(date.today(),'%m%d')+'('+date.strftime(yesterday,'%m%d')+'庫存)'
    ws0.title=t0
    t1='NB(P4A135P)'+date.strftime(date.today(),'%m%d')+'('+date.strftime(yesterday,'%m%d')+'庫存)'
    ws1.title=t1

    #定位空白格數BM,BP
    B=[]
    for row in ws0['B1':'B200']:
        for cell in row:
            if cell.value==None:
                B+=[cell.row]
    BM=min(B)-1
    for i in range(len(B)-1):
        if (B[i+1]-B[i])!=1:
            BP=B[i+1]-1

    #寫入資料
    for x in range(1,len(nameM)+1):
        ws0['A'+str(x+BM)]=x
        ws0['B'+str(x+BM)]=date.strftime(date.today(),'%Y/%m/%d')
        ws0['C'+str(x+BM)]=date.strftime(yesterday,'%Y/%m/%d')
        ws0['D'+str(x+BM)]='高雄長庚紀念醫院'
        ws0['E'+str(x+BM)]='P4A151M'
        ws0['F'+str(x+BM)]='Molunpiravir'
        ws0['G'+str(x+BM)]='W005595'
        ws0['H'+str(x+BM)]=nameM[x-1]
        ws0['I'+str(x+BM)]=IDM[x-1]
        ws0['J'+str(x+BM)]='NB'
    ws0['K'+str(len(nameM)+BM)]=len(nameM)
    for x in range(1,len(nameP)+1):
        ws0['A'+str(x+BP)]=x
        ws0['B'+str(x+BP)]=date.strftime(date.today(),'%Y/%m/%d')
        ws0['C'+str(x+BP)]=date.strftime(yesterday,'%Y/%m/%d')
        ws0['D'+str(x+BP)]='高雄長庚紀念醫院'
        ws0['E'+str(x+BP)]='P4A152M'
        ws0['F'+str(x+BP)]='Paxlovid'
        ws0['G'+str(x+BP)]='GC4219'
        ws0['H'+str(x+BP)]=nameP[x-1]
        ws0['I'+str(x+BP)]=IDP[x-1]
        ws0['J'+str(x+BP)]='NB'
    ws0['K'+str(len(nameP)+BP)]=len(nameP)
    wb.save(filename)
except:
    with open('發生錯誤.txt','w',encoding='utf-8') as Res:
        print('執行失敗!\n可能是找不到使用清單檔案,當日xlsx檔\n或者是表格已滿,請先添加空格',file=Res)