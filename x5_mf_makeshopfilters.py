import sys
import os
import re
import xlrd

from tkinter import filedialog as tkd

file_path_string = tkd.askopenfilename()

book = xlrd.open_workbook(file_path_string)
#print(book.sheet_names()[1])

def readsheet(sheet):
    numrows = sheet.nrows
    numcols = sheet.ncols
    tb=[]
    markers = []
    for i in range(numrows):
        tr = []
        for j in range(numcols):
            cell_value = sheet.cell(i,j).value
            try:
                tr.append(int(cell_value))
            except:
                try:
                    tr.append(float(cell_value))
                except:
                    try:
                        tr.append(str(cell_value))
                        if 'Код' in cell_value:
                            markers.append({'row':i,'col':j,'value':cell_value})                        
                    except:
                        tr.append('')
        #print(tr)
        tb.append(tr)
    return tb, markers, numcols, numrows

itercount = 0
neededLists = ['Межформ конкуренты','Список онлайн-конкурентов']

gen = [x for x in book.sheet_names() if x in neededLists]

for listname in gen:
    tb, markers, numcols, numrows = readsheet(book.sheet_by_name(listname))
    print(markers)
    for marker in markers:
        if any(x in marker['value'] for x in ['→','->']):
            checkrow = int(marker['row'])
            iterablecolstart = int(marker['col'])+1
        if '↓' in marker['value']:
            checkcol = int(marker['col'])
            iterablerowstart = int(marker['row'])+1

    print(checkrow, checkcol)
    print(iterablerowstart)
    citieslist = [{'code':tb[checkrow][kx],'label':tb[checkrow+2][kx]} for kx in range(iterablecolstart,numcols)]
    extracities_labels_str = 'Ярославль, Рязань, Новосибирск, Великий Новгород, Оренбург, Краснодар, Курск, Ижевск, Чебоксары, Владимир, Тула, Киров, Челябинск, Магнитогорск, Омск, Пермь, Сызрань, Волгоград, Старый Оскол, Тамбов, Белгород, Набережные Челны, Тверь, Кострома, Смоленск'
    extracities_codes_str = ','.join(['_'+str(ky['code']) for ky in citieslist if ky['label'] in [lbl.strip() for lbl in extracities_labels_str.split(',')]])

    #print(citieslist)
    #print(extracities_codes_str)

    filtshops = []
    for j in range(iterablecolstart, numcols):
        filtshop = []
        for i in range(iterablerowstart, numrows):
            if len(str(tb[i][j]))>0:
                filtshop.append(tb[i][checkcol])
        filtshops.append({'city':tb[checkrow][j],'filter':filtshop})
        #print(tb[checkrow][j],' ',filtshop)

    with open(file_path_string[:file_path_string.rfind('/')+1]+"output "+ listname +".txt", "w") as f:
        for qitem in filtshops:
            q_output = "case {_"+str(qitem['city'])+"}\n\tfiltershop"+str(itercount)+"={"+','.join(["_" + str(y) for y in qitem['filter']])+"}\n"
            f.write(q_output)
        f.write(extracities_codes_str)
    itercount +=1
    