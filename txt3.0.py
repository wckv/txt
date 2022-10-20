try:
    with open('公費克流感使用清單.txt','r',encoding='big5') as input:
        input=input.read()
        input=input[101:]
        n=int(input.count('#')/19)
        input=input.split('#')
    name=input[1:19*n:19]
    name=','.join(name)
    name+=','
    name=name.replace(',','\t,')
    name=name.split(sep=',')
    del name[-1]

    drug=input[8:19*n:19]

    ID=input[4:19*n:19]
    ID=','.join(ID)
    ID=ID.replace(',','\n,')
    ID=ID.split(sep=',')

    listM=[]
    listP=[]

    for (a,b) in list(enumerate(drug)):
        if b=='P4A151M':
            listM+=[a]
        elif b=='P4A152M':
            listP+=[a]

    nameM=[]
    IDM=[]
    nameP=[]
    IDP=[]

    for a in listM:
        nameM+=name[a:a+1]
        IDM+=ID[a:a+1]
    for a in listP:
        nameP+=name[a:a+1]
        IDP+=ID[a:a+1]

    from itertools import chain
    outputM=list(chain.from_iterable(zip(nameM, IDM)))
    outputM=''.join(outputM)
    outputP=list(chain.from_iterable(zip(nameP, IDP)))
    outputP=''.join(outputP)

    with open('EXCEL.txt','w',encoding='utf-8') as result:
        print(outputM,file=result,end='\n'*3)
        print(outputP,file=result)

except:
    with open('執行失敗.txt','w',encoding='utf-8') as result:
        print('無法讀取檔案',file=result)
        print('可能沒有文字檔案存在',file=result)
        print('或者姓名可能有罕字無法運作，把罕字刪掉再來',file=result)
