from tkinter import *
from tkinter.ttk import *
from tkinter import filedialog as fd
import pylightxl.pylightxl.pylightxl as xl


def callback(): #rename from callback bcuz then im gonna have to save the file too in a different method #or not lol
    global inPath;
    inPath= fd.askopenfilename()
    path_text.insert('end', inPath)
    
def importCallback(): #rename from callback bcuz then im gonna have to save the file too in a different method #or not lol
    global templatePath;
    templatePath= fd.askopenfilename()
    template_text.insert('end', templatePath)
    #put values into template_cb
    dbTemplate = xl.readxl(fn=templatePath)
    template_cb['values']= dbTemplate.ws_names


# def importExcel():
#     print("h")

def export():
    db = xl.readxl(fn=inPath) #db with input data #delete first(actually second row) because its not needed and its messing things up

    

    print(inPath)
    print("hello")
    #creaty empty dbs
    dbOut = xl.Database()
    dbOut.add_ws(ws="Sheet1") 
    #have to add EANs etc before all this 
    #different for VN and NN though
    if(current_var_1.get()=='VN'):
        dbTemplate = xl.readxl(fn=templatePath)
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
        
        for i, cell in enumerate(dbTemplate.ws(ws=dbTemplate.ws[current_var.get()]).col(col=2), start=1):
            dbOut.ws(ws="Sheet1").update_index(row=i, col=1, val=cell)


        #inCols = [41,20,160,1, 142,104,86,83,106] #41 is EAN - copying that from template, then matching with this
        inCols = [20,160,1, 142,104,86,83,106] 

        for i, ean in enumerate(dbOut.ws(ws="Sheet1").col(col=1), start=1):
            if not(i>1 and ean != dbOut.ws(ws='Sheet1').index(row=i-1, col=1)):
                continue
            for j, cell in enumerate(db.ws(ws="Sheet1").col(col=41), start=1):
                if cell != ean:
                    continue
                #write whole row
                for k, item in enumerate(inCols, start=1):
                    dbOut.ws(ws="Sheet1").update_index(row=i, col=k+1, val=(db.ws(ws="Sheet1").index(row=j, col=item)))
                break #break from for j??


        # for i,item in enumerate(inCols, start=2):
        #     for j, cell in enumerate(db.ws(ws='Sheet1').col(col=item), start=1):
        #         #if col41 rowj == dbout col1 rowj #db.ws(ws='Sheet1').index(row=1, col=2)
        #         if(db.ws(ws='Sheet1').index(row=j, col=41)==dbOut.ws(ws='Sheet1').index(row=j, col=1)): #this is bs
        #             dbOut.ws(ws="Sheet1").update_index(row=j, col=i, val=cell)

        
        #for each line in dbout, when not empty, add newcol and newcolline++

        newLine=1

        for i, item in enumerate(dbOut.ws(ws="Sheet1").col(col=4), start=1): #only if theres value in the line
            if len(item)>0:
                dbOut.ws(ws="Sheet1").update_index(row=i, col=4, val=newCol[newLine+1])
                dbOut.ws(ws="Sheet1").update_index(row=i, col=10, val="=C{}+D{}+E{}+F{}+G{}+H{}+I{}".format(i,i,i,i,i,i,i))
                newLine+=1
                
            
    #also need to edit this oh god
    elif(current_var_1.get()=='NN'):
        secondCol = [177,180,174,95,88]
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

        inCols = [41,20,112,1,160]

        for i,item in enumerate(inCols, start=1):
            for j, cell in enumerate(db.ws(ws='Sheet1').col(col=item), start=1):
                dbOut.ws(ws="Sheet1").update_index(row=j, col=i, val=cell)
        for i, item in enumerate(newCol, start=1):
            dbOut.ws(ws="Sheet1").update_index(row=i, col=4, val=newCol[i-1])
            dbOut.ws(ws="Sheet1").update_index(row=i, col=6, val="=C{}+D{}+E{}".format(i,i,i))
        #sums of all as last col
            
    # else:
    #     #error, have to set VN as default and add message box to display error
    #     #https://docs.python.org/3/library/tkinter.messagebox.html
    #     print("coze")
    outName=fd.asksaveasfilename(initialfile="output", defaultextension=".xlsx")
    xl.writexl(db=dbOut, fn=outName)

root = Tk()
root.geometry('350x300')
root.title(':3')

current_var = StringVar()

template_label = Label(root, text="Template")
template_text = Text(root, height = 2, width = 10) #make readonly
import_btn = Button(root, text="Import", command=importCallback)
template_cb=Combobox(root, textvariable=current_var, width=10)


source_label=Label(root, text="Source")
path_text = Text(root, height = 2, width = 25) #make readonly
browse_btn = Button(root, text="Browse", command=callback)


current_var_1 = StringVar()
combobox = Combobox(root, textvariable=current_var_1, width=25) 
combobox['values']=('VN', 'NN')
combobox['state'] = 'readonly'
combobox.set("VN")

export_btn = Button(root, text="Export", command=export)

#https://www.pythontutorial.net/tkinter/tkinter-progressbar/

output_text = Text(root, height = 4, width = 40) #also make readonly - will be used for errors or sth maybe it might not be necessarry 


template_label.grid(column=0, row=0, padx=10, pady=10)
template_text.grid(column=0, row=1,)
import_btn.grid(column=2, row=1)
template_cb.grid(column=1, row=1)

source_label.grid(column=0, row=2, padx=10, pady=10)
path_text.grid(column=0, row=3,columnspan=2)
browse_btn.grid(column=2, row=3)
combobox.grid(column=0, row=4,columnspan=2, padx=10, pady=10)
export_btn.grid(column=2, row=4)
output_text.grid(column=0, row=5, columnspan=3, padx=10)

#https://pylightxl.readthedocs.io/en/latest/quickstart.html

root.mainloop()