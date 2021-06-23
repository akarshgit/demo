# Python program to create
# a file explorer in Tkinter
  
# import all components
# from the tkinter library

# narecop302@moxkid.com
# kahaselau
# TSM%],z=3jy"]7Wu

from tkinter import *
from tkinter.ttk import Progressbar
from tkinter import ttk
import time
  
# import filedialog module
from tkinter import filedialog

# Circular Progress
# https://stackoverflow.com/questions/41943823/circular-progress-bar-using-tkinter

def perform_operation(filename):
    print(filename)
    with open('test.txt', 'w') as f:
        f.write(filename)
    if 'xlsx' in filename:
        label_status.configure(text="Updated", fg='green')
    else:
        label_status.configure(text="Not Updated", fg='red')

    for i in range(10):
        bar['value'] += 10
        print(i)
        time.sleep(1)

  
# Function for opening the
# file explorer window
def browseFiles():
    filename = filedialog.askopenfilename(initialdir = "/",
                                          title = "Select a File",
                                          filetypes = (("Excel files",
                                                        "*.xlsx*"),
                                                       ("All files",
                                                        "*.*")))
      
    # Change label contents
    # label_file_explorer.configure(text="File Opened: "+filename)
    
    perform_operation(filename)
      

                                                                                                  
# Create the root window
window = Tk()
  
# Set window title
window.title('Database Updater')
  
# Set window size
window.geometry("500x500")
window.resizable(0, 0)
  
#Set window background color
window.configure(background='#101417')

window.pack_propagate(False)


logo = PhotoImage(file = "logo.png")
logo_btn = Button(window,
                        image = logo, borderwidth=0,
                        height = 81, width = 86, bg="#101417")


s = ttk.Style()
s.theme_use('winnative')
s.configure("black.Horizontal.TProgressbar", background='#0563bb', bordercolor="red")
bar = Progressbar(window, length=100)#, style='black.Horizontal.TProgressbar')
bar['value'] = 0

# Create a File Explorer label
label_file_explorer = Label(window,
                            text = "Update Excel Metadata in Sharepoint",
                            height = 2,
                            fg = "#0563bb", bg="#101417",
                            font=("Open Sans", 20))

label_status = Label(window,
                            text = "",
                            width = 71, height = 1,
                            fg = "#0563bb", bg="#101417")

label_ask_file = Label(window,
                            text = "Please Select an Excel File",
                            height = 2,
                            fg = "#728394", bg="#101417",
                            font=("Open Sans", 12))

btn_image = PhotoImage(file = "bb.png")
button_explore = Button(window,
                        text = "Browse Files",
                        command = browseFiles,
                        image = btn_image, borderwidth=0,
                        height = 37, width = 120, bg="#101417")

#button_explore.config(background="#ffffff", borderwidth=0)

  
#b2 = button_exit = Button(window,
#                     text = "Exit",
#                     command = exit)
#b2.grid(row=0, column=0, pady=(10, 10))
  
# Grid method is chosen for placing
# the widgets at respective positions
# in a table like structure by
# specifying rows and columns

logo_btn.grid(column = 1, row = 1, pady=(5,0))

label_file_explorer.grid(column = 1, row = 2)
  
button_explore.grid(row=4, column=1, pady=(20, 20))
  
#button_exit.grid(column = 1,row = 3)

bar.grid(row=5, column=1, pady=(40, 0))
  
label_status.grid(column = 1,row = 6, pady=(80, 0))

label_ask_file.grid(column = 1, row=3)

label_status.config(text = "Hello World", fg="green")

#label_status.place(relx=1.0, rely=1.0, anchor='center')

# Let the window wait for any events
window.mainloop()