with open('M.txt',encoding='utf-8') as M:
    M=M.read()
    M=M.replace(' ','')
    M=M.replace('\n','\t')
    ListM=M.split('\t')
    nM=int(len(ListM))-2

with open('P.txt',encoding='utf-8') as P:
    P=P.read()
    P=P.replace(' ','')
    P=P.replace('\n','\t')
    ListP=P.split('\t')
    nP=int(len(ListP))-2

i=0
strM=''
while i<nM:
    p='【'+ListM[i]+' '+ListM[i+1]+' '+ListM[i+2]+'】'
    strM=strM+p
    i+=3

j=0
strP=''
while j<nP:
    p='【'+ListP[i]+' '+ListP[i+1]+' '+ListP[i+2]+'】'
    strP=strP+p
    j+=3

import datetime
today=datetime.date.today()
yesterday=today+datetime.timedelta(-1)
today=today.strftime('%Y/%m/%d')
yesterday=yesterday.strftime('%m/%d')

result=open('result.txt','w',encoding='utf-8')

nM=int((nM+2)/3)
M=today+'申報'+yesterday+'處方【Molnupiravir批號W005595=>'+str(nM)+'PC】'+strM
print(M,'\n'*3,file=result)

nP=int((nP+2)/3)
P=today+'申報'+yesterday+'處方【Paxlovid批號GC4219=>'+str(nP)+'PC】'+strP
print(P,file=result)

result.close()

