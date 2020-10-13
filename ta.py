import sys
import os
import re
import xlrd
import itertools


from tkinter import filedialog as tkd

file_path_string = tkd.askopenfilename()

book = xlrd.open_workbook(file_path_string)
sheets = book.sheet_names()

def readbycol(sheet):
    databycol = []
    for c_ind in range(sheet.ncols):
        #databycol.append(sheet.col_values(c_ind))
        spl = [list(y) for x, y in itertools.groupby(sheet.col_values(c_ind), lambda z: z == '#') if not x]
        databycol.append(dict(zip(['code','ta1','ta2','ta3'],spl)))

    return databycol

data = readbycol(book.sheet_by_name(sheets[0])) 
fulldata = [readbycol(book.sheet_by_name(sh_elem)) for sh_elem in sheets]

print(data)

with open(file_path_string[:file_path_string.rfind('/')+1]+"output_ta.txt", "w") as f:
    for i in range(len(sheets)):
        for item in fulldata[i]:
            q_output = "case {"+','.join(["_"+str(int(a)) for a in item['code']])+"}\n\tta1={"+','.join(["_" + str(int(b)) for b in item['ta1'] if b!=''])+"}\n\tta2={"+','.join(["_" + str(int(c)) for c in item['ta2'] if c!=''])+"}\n\tta3={"+','.join(["_" + str(int(d)) for d in item['ta3'] if d!=''])+"}\n"
            f.write(q_output)
        f.write("\n\n")


