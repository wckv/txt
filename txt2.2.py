with open('input.txt','r',errors='ignore',encoding='utf-8') as In:
    lst=In.read()
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
with open('result.txt','w',errors='ignore',encoding='utf-8') as Res:
    import datetime
    today=datetime.date.today()
    yesterday=today+datetime.timedelta(-1)
    today=today.strftime('%Y/%m/%d')
    yesterday=yesterday.strftime('%m/%d')
    date=today+'申報'+yesterday+'處方'
    print(date,'【Molnupiravir批號W005595=>',nM,'PC】',sep='',end='',file=Res)
    print(''.join(M),sep='',end='\n'*3,file=Res)
    print(date,'【Paxlovid批號GC4219=>',nP,'PC】',sep='',end='',file=Res)
    print(''.join(P),sep='',end='',file=Res)