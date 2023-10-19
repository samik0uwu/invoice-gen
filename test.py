from tkinter import filedialog as fd
import pylightxl.pylightxl.pylightxl as xl


db = xl.readxl(fn='inputVN.xlsx')


data =db.ws(ws='Sheet1').col(col=1)

# for row_id, item in enumerate(data, start=1):
#     print(item)
#     print(row_id)


# for item in enumerate(data, start=1):
#     print(item[1])


secondCol = [112,71,77,88,95]

dbs = [];
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
    newCol.append(x)

for item in newCol:
    print("{:.2f}".format(item))