from tkinter import filedialog as fd
import pylightxl.pylightxl.pylightxl as xl


db = xl.readxl(fn='inputVN.xlsx')
data =db.ws(ws='Sheet1').col(col=1)#?? what even is this :sob:

    
dbOut = xl.Database()#creaty empty dbs
dbOut.add_ws(ws="Sheet1") 

secondCol = [112,71,77,88,95]
dbs = []; #db with all cols from secondCol
for i,item in enumerate(secondCol, start=0):
    dbs.append(db.ws(ws='Sheet1').col(col=item))
newCol = [];
for i in range(len(dbs[0])):
    x=0
    for j in range(len(dbs)):
        if (isinstance(dbs[j][i], float) or isinstance(dbs[j][i], int)):
            x+=dbs[j][i]
        else:
            break;        
    newCol.append(x) #all into one list with one col

inCols = [160,1, 142,104,86,83,106]

for i,item in enumerate(inCols, start=1):
    for j, cell in enumerate(db.ws(ws='Sheet1').col(col=item), start=1):
        dbOut.ws(ws="Sheet1").update_index(row=j, col=i, val=cell)

    #dataOut[i] =db.ws(ws='Sheet1').col(col=item)
#dataOut[1]=newCol
for i, item in enumerate(newCol, start=1):
    dbOut.ws(ws="Sheet1").update_index(row=i, col=2, val=newCol[i-1])


# for i in range(7):
#     for row_id, item in enumerate(data[i], start=1): #i need to create data before the for i think?
#         dbOut.ws(ws="Sheet1").update_index(row=row_id, col=i, val=item) #dont use outcol, figure it out lol

name = "outputVN.xlsx"
xl.writexl(db=dbOut, fn=name)




