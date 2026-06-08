import tkinter as tk
from tkinter import *
root = tk.Tk()

f = tk.Frame(root)
f.place(x=10, y=20)

scrollbar = Scrollbar(f)
t = tk.Text(f, height=10, width=10, yscrollcommand=scrollbar.set)
scrollbar.config(command=t.yview)
scrollbar.pack(side=RIGHT, fill=Y)
t.pack(side="left")

root.mainloop()
