with open('公費克流感使用清單.txt','r',encoding='big5',errors='ignore') as input:
    input=input.read()[101:]
    n=int(input.count('#')/19)
    input=input.split('#')
    name=input[1:19*n:19]
    ID=input[4:19*n:19]
    drug=input[8:19*n:19]
    lst=list(zip(name,ID,drug))
with open('執行結果.txt','w',encoding='utf-8',errors='ignore') as Res:
    print('Molnupiravir,P4A151M',file=Res)
    for i in range(len(lst)):
        if 'P4A151M' in lst[i]:
            print('\t'.join(lst[i][0:2]),file=Res)
    print('\n'*3,'Paxlovid,P4A152M', sep='',file=Res)
    for i in range(len(lst)):
        if 'P4A152M' in lst[i]:
            print('\t'.join(lst[i][0:2]),file=Res)