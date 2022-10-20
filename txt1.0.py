with open('1.txt',encoding='utf-8') as str:
    str=str.read()
    str=str.replace(' ','')
    str=str.replace('\n','\t')
    lst=str.split('\t')
    n=int(len(lst))-2
i=0
str2=''
result=open('2.txt','w',encoding='utf-8')
while i<n:
    p='【'+lst[i]+' '+lst[i+1]+' '+lst[i+2]+'】'
    str2=str2+p
    i+=3
import datetime
today=datetime.date.today()
yesterday=today+datetime.timedelta(-1)
today=today.strftime('%Y/%m/%d')
yesterday=yesterday.strftime('%m/%d')
date=today+'申報'+yesterday+'處方'
print(date,file=result, end='')
n=int((n+2)/3)
print('【 批號 =>',n,'PC】',file=result, end='')
print(str2,file=result)
result.close()