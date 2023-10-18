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

def writeCol(inPath, outPath, inCol, outCol):
    db = xl.readxl(fn=inPath)
    #db.ws(ws='Sheet1').range('AO2:AO48')
    data =db.ws(ws='Sheet1').col(col=inCol)
    dbOut = xl.Database()
    dbOut.add_ws(ws="Sheet1")
    if(current_var.get=='VN'):
        generateVN(data, dbOut)
    for row_id, item in enumerate(data, start=1):
        dbOut.ws(ws="Sheet1").update_index(row=row_id, col=outCol, val=item)
    xl.writexl(db=dbOut, fn=outPath)

def generateVN(data, db):
    for row_id, item in enumerate(data, start=1):
        db.ws(ws="Sheet1").update_index(row=row_id, col=1, val=item)




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

#make a method : input col and output col int, i just get the col from input file and put it into output except the ones i have to add tgether

# db = xl.readxl(fn='inputVN.xlsx')


# db.ws(ws='Sheet1').range('AO2:AO48')


# data =db.ws(ws='Sheet1').col(col=41)

# # create a blank db
# dbOut = xl.Database()
# # add a blank worksheet to the db
# dbOut.add_ws(ws="Sheet1")


# for item in enumerate(data, start=1):
#     dbOut.ws(ws="Sheet1").update_index(row=1, col=1, val=item)

# #dbOut.ws(ws="Sheet1").update_index(row=1, col=1, val=data)

# xl.writexl(db=dbOut, fn="output.xlsx")



root.mainloop()