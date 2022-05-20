from tkinter import filedialog

typ = [('テキストファイル','*.png')] 
dir = 'C:\\pg'
fle = filedialog.askopenfilename(filetypes = typ, initialdir = dir) 

print(fle)