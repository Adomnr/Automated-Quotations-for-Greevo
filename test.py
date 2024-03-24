import tkinter
from tkinter import ttk



def hide_button():
    user_info_frame.grid_remove()

def show_button():
    user_info_frame.grid()

def seletioncombobox(*args):
    if newcombobox.get() == '1':
        hide_button()
    else:
        show_button()


window = tkinter.Tk()
frame = tkinter.Frame(window)
frame.pack()

# Saving User Info
user_info_frame = tkinter.LabelFrame(frame, text="Panel Inverters and PV Balance")
user_info_frame.grid(row=0, column=0, padx=20, pady=10)

button = tkinter.Button(frame, text="Enter data")
button.grid(row=1, column=0, sticky="news", padx=20, pady=10)

user_info_frame = tkinter.LabelFrame(frame, text="Panel Inverters and PV Balance")
user_info_frame.grid(row=0, column=0, padx=20, pady=10)

button = tkinter.Button(user_info_frame, text="Button 1")
button.grid(row=2, column=0, sticky="news", padx=20, pady=10)

button = tkinter.Button(user_info_frame, text="Button 2")
button.grid(row=4, column=0, sticky="news", padx=20, pady=10)

button = tkinter.Button(frame, text="Button 2", command=show_button)
button.grid(row=5, column=0, sticky="news", padx=20, pady=10)

button = tkinter.Button(frame, text="Button 2", command=hide_button)
button.grid(row=6, column=0, sticky="news", padx=20, pady=10)

hider_frame = tkinter.LabelFrame(frame, text="Nah")
hider_frame.grid(row=6, column=0, padx=20, pady=10)

sel = tkinter.StringVar()

newcombobox = ttk.Combobox(hider_frame,values=['1','2'], textvariable=sel)
newcombobox.grid(row=9,column=0)

sel.trace('w',seletioncombobox)

window.mainloop()