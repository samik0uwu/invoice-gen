from tkinter import *
from tkinter.ttk import *
from tkinter import filedialog as fd
import pylightxl.pylightxl.pylightxl as xl


def callback(): #rename from callback bcuz then im gonna have to save the file too in a different method #or not lol
    global inPath;
    inPath= fd.askopenfilename()
    path_text.insert('end', inPath)

def export():
    db = xl.readxl(fn=inPath) #db with input data
    print(inPath)
    print("hello")
    #creaty empty dbs
    dbOut = xl.Database()
    dbOut.add_ws(ws="Sheet1") 
    #have to add EANs etc before all this
    if(current_var.get()=='VN'):
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

        inCols = [160,1, 142,104,86,83,106]

        for i,item in enumerate(inCols, start=1):
            for j, cell in enumerate(db.ws(ws='Sheet1').col(col=item), start=1):
                dbOut.ws(ws="Sheet1").update_index(row=j, col=i, val=cell)
        for i, item in enumerate(newCol, start=1):
            dbOut.ws(ws="Sheet1").update_index(row=i, col=2, val=newCol[i-1])

    elif(current_var.get()=='NN'):
        print("hi")
    else:
        #error, have to set VN as default and add message box to display error
        #https://docs.python.org/3/library/tkinter.messagebox.html
        print("coze")
    outName=fd.asksaveasfilename(initialfile="outputVN", defaultextension=".xlsx")
    xl.writexl(db=dbOut, fn=outName)

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
combobox.set("VN")

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