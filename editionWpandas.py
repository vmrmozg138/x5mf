import pandas as pd
from tkinter import filedialog as tkd

fPathStr = tkd.askopenfilename()
fname = fPathStr[fPathStr.rfind('/'):fPathStr.rfind('.')]
pathPrefix = fPathStr[:fPathStr.rfind('/')]
print(fPathStr,'\n',fname,'\n',pathPrefix)

def is_number(s):
    try:
        float(s)
        return True
    except ValueError:
        return False

if fPathStr.endswith('xlsx'):
    xl = pd.ExcelFile(fPathStr)
elif fPathStr.endswith('csv'):
    data = pd.read_csv(fPathStr,encoding='utf-16',delimiter='\t')

basename = {'Короткая ссылка':'base_short.xlsx','Длинная ссылка':'base_long.xlsx'}

if xl:
    sheetNames = xl.sheet_names
    for sn in sheetNames:
        print(sn)
        print(basename[sn])
        
        data = pd.read_excel(xl, sn)
        base = pd.read_excel(pathPrefix + '/'+basename[sn])
        df = pd.DataFrame(data) #текущая вкладка
        bdf = pd.DataFrame(base) #соответствующая ей база
        print(bdf)
        colnames = list(df.columns)
        bcolnames = list(bdf.columns)
        print(colnames,'\n',bcolnames)
        resdf = pd.DataFrame(columns=bcolnames)
        for index,row in df.iterrows():
            listrow = list(row)
            for ind,datacell in enumerate(listrow[1:]):
                if str(datacell).strip()!='nan':
                    excess = abs(int(datacell))  # получаем число лишних из ячейки
                    city = str(listrow[0]).strip()            # получение города
                    supplier = str(colnames[ind+1]).strip()     # получение подрядчика - +1 потому что индекс на 1 меньше
                    print(f'город - {str(city)}, подрядчик - {str(supplier)}, лишних людей - {str(excess)}') 
                    # теперь выберем подходящих людей
                    varCity = '{'+ str(city) + '}'
                    varSupplier = '{'+ str(supplier) + '}'
                    
                    people = bdf[(bdf[' QRecodeCity'] == varCity) & (bdf[' supplierid'] == varSupplier)]
                    #print(people)
                    target_people = people.tail(excess)
                    print(target_people)
                    resdf = resdf.append(target_people)
                    print(resdf)

                    resdf.to_csv(pathPrefix+'/'+sn+'.csv',sep='\t')

        



