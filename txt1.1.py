with open('1.txt',encoding='utf-8') as input:
    input=input.read()
    input=input.split()
n=len(input)
output=''
i=0
while i<n:
    j=input[i:i+3]
    j=' '.join(j)
    output+='【'+j+'】'
    i+=3
result=open('2.txt','w',encoding='utf-8')
import datetime
today=datetime.date.today()
yesterday=today+datetime.timedelta(-1)
today=today.strftime('%Y/%m/%d')
yesterday=yesterday.strftime('%m/%d')
date=today+'申報'+yesterday+'處方'
print(date,file=result, end='')
print('【 批號 =>',int(n/3),'PC】',file=result, end='')
print(output,file=result)
result.close()