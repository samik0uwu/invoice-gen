from tkinter import *
from tkinter.ttk import *
from tkinter import filedialog as fd


def callback(): #rename from callback bcuz then im gonna have to save the file too in a different method
    name = fd.askopenfilename()
    path_text.insert('end', name)


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

export_btn = Button(root, text="Export")

#https://www.pythontutorial.net/tkinter/tkinter-progressbar/



output_text = Text(root, height = 4, width = 40) #also make readonly - will be used for errors or sth maybe it might not be necessarry 





path_text.grid(column=0, row=0, padx=10, pady=10)
browse_btn.grid(column=1, row=0)
combobox.grid(column=0, row=1)
export_btn.grid(column=1, row=1)
output_text.grid(column=0, row=3, columnspan=2, padx=10, pady=10)


root.mainloop()