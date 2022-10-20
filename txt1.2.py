with open('1.txt','r',encoding='utf-8') as input:
    str=input.read()
    str=str.split()
    n=len(str)
    import datetime
    today=datetime.date.today()
    yesterday=today+datetime.timedelta(-1)
    today=today.strftime('%Y/%m/%d')
    yesterday=yesterday.strftime('%m/%d')
    date=today+'申報'+yesterday+'處方'
with open('2.txt','w',encoding='utf-8') as output:
    print(date,'【 =>',int(n/3),'PC】',sep='',end='',file=output)
    i=0
    while i<n:
        j=str[i:i+3]
        j=' '.join(j)
        print('【',j,'】',sep='',end='',file=output)
        i+=3