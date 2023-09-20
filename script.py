from tkinter import *
from tkinter.ttk import *
from tkinter import filedialog as fd

def callback():
    name = fd.askopenfilename()
    path_text.insert('end', name)


root = Tk()
root.geometry('350x200')
root.title(':3')

path_text = Text(root, height = 2, width = 20)
browse_btn = Button(root, text="Browse", command=callback)

#https://www.pythontutorial.net/tkinter/tkinter-combobox/
current_var = StringVar()
combobox = Combobox(root, textvariable=current_var)
combobox['values']=('VN', 'NN')
combobox['state'] = 'readonly'

export_btn = Button(root, text="Export")
output_text = Text(root, height = 4, width = 40)



path_text.grid(column=0, row=0, padx=10, pady=10)
browse_btn.grid(column=1, row=0)
combobox.grid(column=0, row=1)
export_btn.grid(column=1, row=1)
output_text.grid(column=0, row=2, columnspan=2, padx=10, pady=10)
#columnspan=2 u output textboxu

root.mainloop()