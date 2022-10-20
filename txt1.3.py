with open('1.txt','r',encoding='utf-8') as input:
    str=input.read()
    str=str.split()
    n=int(len(str)/3)
    import datetime
    today=datetime.date.today()
    yesterday=today+datetime.timedelta(-1)
    today=today.strftime('%Y/%m/%d')
    yesterday=yesterday.strftime('%m/%d')
    date=today+'申報'+yesterday+'處方'
with open('2.txt','w',encoding='utf-8') as output:
    print(date,'【 =>',n,'PC】',sep='',end='',file=output)
    for i in range(n):
        j=' '.join(str[i:i+3])
        print('【',j,'】',sep='',end='',file=output)