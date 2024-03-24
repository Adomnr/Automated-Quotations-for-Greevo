import tkinter
from tkinter import ttk



def hide_button():
    user_info_frame.grid_remove()

def show_button():
    user_info_frame.grid()

def seletioncombobox(*args):
    print("here")
    if newcombobox.get() == '1':
        hide_button()
        print("nowhere")
    else:
        show_button()
        print("nowhere meow")


window = tkinter.Tk()
frame = tkinter.Frame(window)
frame.pack()

hider_frame = tkinter.LabelFrame(frame, text="Nah")
hider_frame.grid(row=0, column=0, padx=20, pady=10)


sel = tkinter.StringVar()

newcombobox = ttk.Combobox(hider_frame,values=['1','2'], textvariable=sel)
newcombobox.grid(row=1,column=0)

sel.trace('w',seletioncombobox)

# Saving User Info
user_info_frame = tkinter.LabelFrame(frame, text="Panel Inverters and PV Balance")
user_info_frame.grid(row=2, column=0,sticky="news", padx=20, pady=10)

sel2 = tkinter.StringVar()

Reffered_label = tkinter.Label(user_info_frame, text="Referred By")
Reffered_combobox = ttk.Combobox(user_info_frame, values=["Madam Rafia", "Engr Sajjad", "Engr Shaban", "Engr Abid", "Engr Ammar Butt","Engr Ubaid", "Sir Nabeel","Engr Osama"], textvariable=sel2)
Reffered_label.grid(row=2, column=1)
Reffered_combobox.grid(row=3, column=1)

sel2.trace('w',seletioncombobox)

button = tkinter.Button(frame, text="Enter data")
button.grid(row=3, column=0, sticky="news", padx=20, pady=10)

button = tkinter.Button(user_info_frame, text="Button 1")
button.grid(row=4, column=0, sticky="news", padx=20, pady=10)

button = tkinter.Button(user_info_frame, text="Button 2")
button.grid(row=5, column=0, sticky="news", padx=20, pady=10)

button = tkinter.Button(frame, text="Button 2", command=show_button)
button.grid(row=6, column=0, sticky="news", padx=20, pady=10)

button = tkinter.Button(frame, text="Button 2", command=hide_button)
button.grid(row=7, column=0, sticky="news", padx=20, pady=10)




window.mainloop()