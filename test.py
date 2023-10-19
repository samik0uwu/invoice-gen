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

# for i, item in enumerate(dbs[0], start=0):
#     print(i)
#     print(item)
#     print("aaaaaaaaaaaaaa")

for i in range(len(dbs[0])):
    x=0
    print("start")
    for j in range(len(dbs)):
        if (isinstance(dbs[j][i], float) or isinstance(dbs[j][i], int)):
            x+=dbs[j][i]
            print("{:.2f}".format(x))
        else:
            print("its not float its "+str(type(dbs[j][i])))

    print("end")
         
    
    print("{:.2f}".format(x))


