from tkinter import *

from tkinter import messagebox


def answer():
    messagebox.showerror("Answer", "Sorry, no answer available")


def callback():
    if messagebox.askyesno('Verify', 'Really quit?'):
        messagebox.showwarning('Yes', 'Not yet implemented')
    else:
        print("quit")
        exit(0)

answer()

#callback()