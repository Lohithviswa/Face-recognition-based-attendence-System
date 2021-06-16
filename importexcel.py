import tkinter as tk
from tkinter import filedialog
import pandas as pd

root= tk.Tk()

canvas1 = tk.Canvas(root, width = 500, height = 500, bg = 'lightsteelblue')
canvas1.pack()

def getExcel ():
    global df
    
    import_file_path = filedialog.askopenfilename()
    df = pd.read_excel (import_file_path)
    #print(roll_numbers)
    print("file opened")
    
roll_numbers=list(df['Roll_numbers'])   
browseButton_Excel = tk.Button(text='Class Roll(Import Excel File)', command=getExcel, bg='green', fg='white', font=('helvetica', 12, 'bold'))
canvas1.create_window(250, 250, window=browseButton_Excel)

root.mainloop()