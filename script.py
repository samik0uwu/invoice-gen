from tkinter import *
from tkinter.ttk import *
from tkinter import filedialog as fd
import pylightxl.pylightxl.pylightxl as xl

def callback(): #rename from callback bcuz then im gonna have to save the file too in a different method #or not lol
    name = fd.askopenfilename()
    path_text.insert('end', name)

def lmao(inPath, outPath, inCol, outCol):
    db = xl.readxl(fn='inputVN.xlsx')
    db.ws(ws='Sheet1').range('AO2:AO48')
    data =db.ws(ws='Sheet1').col(col=41)
    dbOut = xl.Database()
    dbOut.add_ws(ws="Sheet1")
    for row_id, item in enumerate(data, start=1):
        dbOut.ws(ws="Sheet1").update_index(row=row_id, col=1, val=item)
    xl.writexl(db=dbOut, fn="output.xlsx")

def export():
    #select output file
    name = fd.asksaveasfile()
    writeCol(path_text, name)

def writeCol(inPath, outPath, outCol):
    db = xl.readxl(fn=inPath) #db with input data
    #creaty empty dbs
    dbOut = xl.Database()
    dbOut.add_ws(ws="Sheet1")
    if(current_var.get=='VN'):
        #sum all these cols to one
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

        inCols = [160,0, 142,104,86,83,106]

        for item in inCols:
            data =db.ws(ws='Sheet1').col(col=item)

        for row_id, item in enumerate(data, start=1): #i need to create data before the for i think?
            dbOut.ws(ws="Sheet1").update_index(row=row_id, col=outCol, val=item) #dont use outcol, figure it out lol
    elif(current_var.get=='NN'):
        print("hi")
    else:
        #error, have to set VN as default and add message box to display error
        print("hello")
    xl.writexl(db=dbOut, fn=outPath)






root = Tk()
root.geometry('350x200')
root.title(':3')

path_text = Text(root, height = 2, width = 20) #make readonly
browse_btn = Button(root, text="Browse", command=callback)

#https://www.pythontutorial.net/tkinter/tkinter-combobox/
current_var = StringVar()
combobox = Combobox(root, textvariable=current_var) #make first value set by default
combobox['values']=('VN', 'NN')
combobox['state'] = 'readonly'

export_btn = Button(root, text="Export", command=export)

#https://www.pythontutorial.net/tkinter/tkinter-progressbar/



output_text = Text(root, height = 4, width = 40) #also make readonly - will be used for errors or sth maybe it might not be necessarry 





path_text.grid(column=0, row=0, padx=10, pady=10)
browse_btn.grid(column=1, row=0)
combobox.grid(column=0, row=1)
export_btn.grid(column=1, row=1)
output_text.grid(column=0, row=3, columnspan=2, padx=10, pady=10)

#https://pylightxl.readthedocs.io/en/latest/quickstart.html


root.mainloop()