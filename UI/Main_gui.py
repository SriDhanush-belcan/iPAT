   
from tkinter import *
from tkinter.ttk import Label
from tkinter import Tk
#import tkinter as tk
from tkinter.simpledialog import askinteger
#from tkinter import *
from tkinter import messagebox
from functools import partial
import tkinter as tk
from tkinter import filedialog
import os

master = Tk()
# Set the Geometry
#master.geometry("750x250")
#def print_width():
#   print("The width of Tkinter window:", m5
# aster.winfo_width(10))
#   print("The height of Tkinter window:", master.winfo_height(20))
# Create a Label

# Import module
#from tkinter import *

# Create object
#rmaster = Tk()

# Adjust size
master.geometry("600x400")
#w = Canvas(master, width=40, height=60)
#master = Canvas(master, width=400, height=300)
#master.pack()
# Set window color
#master.configure(bg='blue')
master['background']='#856ff8'

master.title('iPAT')

def printDetails(usernameEntry) :
    usernameText = usernameEntry.get()
    print("user entered :", usernameText)
    return

#Label(master, text='QN#                 :').place(x=5,y=10)
# Entry for user input
#usernameEntry = Entry(master).place(x=100,y=10)
# Label
usernameLabel = Label(master, text='QN#                           :').place(x=5,y=10)
# Entry for user input
usernameEntry = Entry(master).place(x=150,y=10)
# Define callable function with printDetails function and usernameEntry argument
printDetailsCallable = partial(printDetails, usernameEntry)


#Label(master, text='Part No        :').place(x=5,y=40)
usernameLabel = Label(master, text='Part No                       :').place(x=5,y=40)
# Entry for user input
usernameEntry = Entry(master).place(x=150,y=40)
# Define callable function with printDetails function and usernameEntry argument
printDetailsCallable = partial(printDetails, usernameEntry)


#Label(master, text='Assy No        :').place(x=5,y=70)
usernameLabel = Label(master, text='Assy No                       :').place(x=5,y=70)
# Entry for user input
usernameEntry = Entry(master).place(x=150,y=70)
# Define callable function with printDetails function and usernameEntry argument
printDetailsCallable = partial(printDetails, usernameEntry)


#Label(master, text='MQI            :').place(x=5,y=100)
usernameLabel = Label(master, text='MQI                           :').place(x=5,y=100)
# Entry for user input
usernameEntry = Entry(master).place(x=150,y=100)
# Define callable function with printDetails function and usernameEntry argument
printDetailsCallable = partial(printDetails, usernameEntry)


#Label(master, text='Part Name      :').place(x=5,y=130)
usernameLabel = Label(master, text='Part Name                     :').place(x=5,y=130)
# Entry for user input
usernameEntry = Entry(master).place(x=150,y=130)
# Define callable function with printDetails function and usernameEntry argument
printDetailsCallable = partial(printDetails, usernameEntry)


#Label(master, text='Serial No      :').place(x=5,y=160)
usernameLabel = Label(master, text='Serial No                       :').place(x=5,y=160)
# Entry for user input
usernameEntry = Entry(master).place(x=150,y=160)
# Define callable function with printDetails function and usernameEntry argument
printDetailsCallable = partial(printDetails, usernameEntry)


#Label(master, text='Vendor         :').place(x=5,y=190)
usernameLabel = Label(master, text='Vendor                           :').place(x=5,y=190)
# Entry for user input
usernameEntry = Entry(master).place(x=150,y=190)
# Define callable function with printDetails function and usernameEntry argument
printDetailsCallable = partial(printDetails, usernameEntry)



Label(master, text='SAP Package PPT:').place(x=5,y=220)

# Function to update the Listbox with folder contents
def browse_folder():
    folder_path = filedialog.askdirectory()  # Open a folder selection dialog
    if folder_path:
        folder_contents.delete(0, tk.END)  # Clear the Listbox
        for item in os.listdir(folder_path):
            folder_contents.insert(tk.END, item)  # Insert folder contents into Listbox

# Create a button to browse for a folder
browse_button = tk.Button(master, text="Browse SAP Package", command=browse_folder).place(x=150,y=220)
#browse_button.pack(pady=10)



Label(master, text='Vendor info PPT:').place(x=5,y=250)
# Function to update the Listbox with folder contents
def browse_folder():
    folder_path = filedialog.askdirectory()  # Open a folder selection dialog
    if folder_path:
        folder_contents.delete(0, tk.END)  # Clear the Listbox
        for item in os.listdir(folder_path):
            folder_contents.insert(tk.END, item)  # Insert folder contents into Listbox

# Create a button to browse for a folder
browse_button2 = tk.Button(master, text="Browse Vendor Info", command=browse_folder).place(x=150,y=250)
#browse_button.pack(pady=10)


Label(master, text='Engine Model   :').place(x=5,y=280)
# Create the list of options 
options_list = ["Option 1", "Option 2", "Option 3", "Option 4"] 
  
# Variable to keep track of the option 
# selected in OptionMenu 
value_inside = tk.StringVar(master) 
  
# Set the default value of the variable 
value_inside.set("Select Engine Model") 
  
# Create the optionmenu widget and passing  
# the options_list and value_inside to it. 
question_menu = tk.OptionMenu(master, value_inside, *options_list).place(x=150,y=280)
#question_menu.pack() 
  
# Function to print the submitted option-- testing purpose 
def print_answers(): 
    print("Selected Option: {}".format(value_inside.get())) 
    return None


Label(master, text='502 Recommends :').place(x=5,y=310)
# Create the list of options 
options_list = ["Option 1", "Option 2", "Option 3", "Option 4"] 
  
# Variable to keep track of the option 
# selected in OptionMenu 
value_inside = tk.StringVar(master) 
  
# Set the default value of the variable 
value_inside.set("Select 502 Recommends Model") 
  
# Create the optionmenu widget and passing  
# the options_list and value_inside to it. 
question_menu = tk.OptionMenu(master, value_inside, *options_list).place(x=150,y=310)
#question_menu.pack() 
  
# Function to print the submitted option-- testing purpose 
def print_answers(): 
    print("Selected Option: {}".format(value_inside.get())) 
    return None

#Label(master, text='Repair         ').grid(row=11)
Label(master, text='Repair         :').place(x=5,y=340)
# Create the list of options 
options_list = ["Option 1", "Option 2", "Option 3", "Option 4"] 
  
# Variable to keep track of the option 
# selected in OptionMenu 
value_inside = tk.StringVar(master) 
  
# Set the default value of the variable 
value_inside.set("Select Repair Model") 
  
# Create the optionmenu widget and passing  
# the options_list and value_inside to it. 
question_menu = tk.OptionMenu(master, value_inside, *options_list).place(x=150,y=340)
#question_menu.pack() 
  
# Function to print the submitted option-- testing purpose 
def print_answers(): 
    print("Selected Option: {}".format(value_inside.get())) 
    return None


#Label(master, text='Output Path    ').grid(row=12)
Label(master, text='Output Path    :').place(x=5,y=370)

# Function to update the Listbox with folder contents
def browse_folder():
    folder_path = filedialog.askdirectory()  # Open a folder selection dialog
    if folder_path:
        folder_contents.delete(0, tk.END)  # Clear the Listbox
        for item in os.listdir(folder_path):
            folder_contents.insert(tk.END, item)  # Insert folder contents into Listbox

# Create a button to browse for a folder
browse_button3 = tk.Button(master, text="Result Path", command=browse_folder).place(x=150,y=370)
#browse_button.pack(pady=10)

#Label(master, text='OK             ').grid(row=13)
#Label(master, text='Cancel         ').grid(row=14)

def show():
   num = askinteger("Input", "Input an Integer")
   print(num)
   
#B = Button(master, text ="OK", command = show)
B = Button(master, text ="OK",height= 2, width=6)
B.place(x=430,y=350)
#B.grid(row=30,column=40)

def show1():
   num1 = askinteger("Input", "Input an Integer")
   print(num1)
   
#B1 = Button(master, text ="Cancel", command = show1)
B1 = Button(master, text ="Cancel",height= 2, width=6)
B1.place(x=500,y=350)
#B1.grid(row=1,column=0)


# Execute tkinter
master.mainloop()