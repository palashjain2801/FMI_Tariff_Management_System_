
#................This code is developed by Palash Jain......
#..........Updated version of this code can be fetch from 
from tkinter import *
from tkinter import ttk
from tkinter import messagebox
import tkinter as tk
import test1 as test
import concat_Excel as concat_Exl
root = Tk()

root.geometry("1350x768")
root.title("Tariff_Management_System")

output_frame = Frame(root)
input_frame = Frame(root)
# --------------------------------



qp1 = "None"
qp2 = "None"
cmb = ttk.Combobox(input_frame, width="10", values=("NONE","Company","Country","State","District","port","Zip"))
cmb2 = ttk.Combobox(input_frame, width="10", values=("NONE","Country","State","District","port","Zip"))
#cmb = Combobox

class TableDropDown(ttk.Combobox):
    def __init__(self, parent):
        self.current_table = tk.StringVar() # create variable for table
        ttk.Combobox.__init__(self, parent)#  init widget
        # self.config(textvariable = self.current_table, state = "readonly", values = ["Customers", "Pets", "Invoices", "Prices"])
        self.current(0) # index of values for current table
        # self.place(x = 50, y = 100, anchor = "w") # place drop down box 
# CompanyName ,PickupCountry,PickupState,PickupDistrict, PickupPortName,PickupZipCode,DestinationCountry,DestinationState,DestinationDistrict,DestinationPortName,DestinationZipCode, Miles, BaseRate
def checkcmbo():
    global qp1 


    if cmb.get() == "Company":

        print(cmb.get())
        qp1= "lower(CompanyName)"
        #messagebox.showinfo("What user choose", "you choose country" +E1.get())
    elif cmb.get() == "Country":
        print(cmb.get())
        qp1 = "lower(PickupCountry)"
        #messagebox.showinfo("What user choose", "you choose state")

    elif cmb.get() == "State":
        print(cmb.get())
        qp1 = "lower(PickupState)"
        #messagebox.showinfo("What user choose", "you choose state")

    elif cmb.get() == "District":
        print(cmb.get())
        qp1 = "lower(PickupDistrict)"
        print(qp1)
        #messagebox.showinfo("What user choose", "you choose district")

    elif cmb.get() == "port":
        print(cmb.get())
        qp1 = "lower(PickupPortName)"
        #messagebox.showinfo("What user choose", "you choose port")
    elif cmb.get() == "Zip":
        print(cmb.get())
        qp1 = "lower(PickupZipCode)"
        #messagebox.showinfo("What user choose", "you choose zip")
    elif cmb.get() == "NONE":
        print(cmb.get())
        qp1 = "NONE"
        #messagebox.showinfo("nothing to show!", "you have to be choose something")

# DestinationCountry,DestinationState,DestinationDistrict,DestinationPortName,DestinationZipCode, Miles, BaseRate
def checkcmbo2():
    global qp2 

    
    if cmb2.get() == "Country":
        print(cmb2.get())
        qp2 = "lower(DestinationCountry)"
        #messagebox.showinfo("What user choose", "you choose state")

    elif cmb2.get() == "State":
        print(cmb2.get())
        qp2 = "lower(DestinationState)"
        #messagebox.showinfo("What user choose", "you choose state")

    elif cmb2.get() == "District":
        print(cmb2.get())
        qp2 = "lower(DestinationDistrict)"
        #messagebox.showinfo("What user choose", "you choose district")

    elif cmb2.get() == "port":
        print(cmb2.get())
        qp2 = "lower(DestinationPortName)"
        #messagebox.showinfo("What user choose", "you choose port")
    elif cmb2.get() == "Zip":
        print(cmb2.get())
        qp2 = "lower(DestinationZipCode)"
        #messagebox.showinfo("What user choose", "you choose zip")
    elif cmb2.get() == "NONE":
        print(cmb2.get())
        qp2 = "NONE"
        #messagebox.showinfo("nothing to show!", "you have to be choose something")

def callmethord():
    checkcmbo()
    checkcmbo2()
    if(cmb.get()!= "NONE" and cmb2.get() != "NONE"):
        query1 = qp1 +"='" + E1.get().lower() + "'"
        query2 =  qp2 + "='" + E2.get().lower() + "'"
        print('printting pickup',query1)
        print('printing drop',query2)
    elif(cmb.get() == "NONE" and cmb2.get() != "NONE"):
        query1 = "NONE"
        query2 = qp2 + "='" + E2.get().lower() + "'"
        print('printting pickup',query1)
        print('printing drop',query2)
    elif(cmb.get() != "NONE" and cmb2.get() == "NONE"):
        query1 = qp1 +"='" + E1.get().lower() + "'"
        query2 = "NONE"

        

    # a = test.cursor.execute("select * from TestTable1 where ",query1,query2)

    insert(test.Databasework(query1,query2))
   

labelTop = ttk.Label(input_frame, text = "Pick Up Information ")
labelTop.grid(columnspan=2, row=0)
cmb.grid(columnspan=2,row=1)
labelTop = ttk.Label(input_frame, text = "Drop Information ")
labelTop.grid(columnspan=2, row=2)
cmb2.grid(columnspan=2,row=3)

btn = ttk.Button(input_frame, text="Get Value",command= callmethord)
btn.grid(column = 8, row = 1,columnspan = 3)

btn1 = ttk.Button(input_frame, text="Update Database",command= concat_Exl.edit_Excel)
btn1.grid(column = 8, row =3 ,columnspan = 3)
# l1 = Label(input_frame,text="Enter Data")
# l1.pack()

E1 = Entry(input_frame,text="pick up")
E1.grid(column = 4, row = 1, columnspan = 3)
E2 = Entry(input_frame,text="Drop of")
E2.grid(column = 4, row = 3, columnspan = 3)










#   --------------------------------------------------------
# output_frame.pack()
xscrollbar = Scrollbar(output_frame, orient=HORIZONTAL)
xscrollbar.grid(row=1, column=0, sticky=N+S+E+W)

yscrollbar = Scrollbar(output_frame)
yscrollbar.grid(row=0, column=1, sticky=N+S+E+W)

text = Text(output_frame, wrap=NONE, height =35, width = 165,
    xscrollcommand=xscrollbar.set,
    yscrollcommand=yscrollbar.set)
text.grid(row=0, column=0,)
xscrollbar.config(command=text.xview)
yscrollbar.config(command=text.yview)
text.insert(END, '\t\t\t\t\t\t\t\t\t\tEither Provide \n\t\t\t\t\t\t\t\t\tPickup information and Drop information \n\t\t\t\t\t\t\t\t\t\t\tOr\n\t\t\t\t\t\t\t\t\tPickup information and NONE \n\t\t\t\t\t\t\t\t\t\t\tOR\n\t\t\t\t\t\t\t\t\tDrop information and NONE\n\n\t\t\t\t\t\t\tNOTE:Make sure all_data.xlsx file exists in folder named out')

def insert(value):
    text.delete("1.0",END)
    text.insert("2.0",value)

input_frame.grid(row = 0, column = 0)
output_frame.grid(row = 1, column= 0)
root.mainloop()