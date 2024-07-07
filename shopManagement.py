import tkinter
from tkinter import *
from tkinter import ttk
from tkinter import font
from tkinter import messagebox
import pandas as pd
from datetime import datetime

# Define the Excel file paths
products_file = "products.xlsx"
sales_file = "sales.xlsx"

# Function to initialize the Excel files if they don't exist
def initialize_excel_files():
    # Check if the products file exists, if not create it
    try:
        pd.read_excel(products_file)
    except FileNotFoundError:
        df = pd.DataFrame(columns=["date", "prodName", "prodPrice"])
        df.to_excel(products_file, index=False)
    
    # Check if the sales file exists, if not create it
    try:
        pd.read_excel(sales_file)
    except FileNotFoundError:
        df = pd.DataFrame(columns=["custName", "date", "prodName", "qty", "price"])
        df.to_excel(sales_file, index=False)

initialize_excel_files()

# Function to add the product to the Excel file
def prodtoTable():
    pname = prodName.get()
    price = prodPrice.get()
    dt = date.get()
    
    # Create a DataFrame with the new product
    new_product = pd.DataFrame({"date": [dt], "prodName": [pname], "prodPrice": [price]})
    
    try:
        # Append the new product to the existing products file
        df = pd.read_excel(products_file)
        df = df.append(new_product, ignore_index=True)
        df.to_excel(products_file, index=False)
        messagebox.showinfo('Success', "Product added successfully")
    except Exception as e:
        print("The exception is:", e)
        messagebox.showinfo("Error", "Trouble adding data into Excel file")
    
    wn.destroy()

# Function to get details of the product to be added
def addProd():
    global prodName, prodPrice, date, Canvas1, wn
    
    wn = tkinter.Tk()
    wn.title("PythonGeeks Shop Management System")
    wn.configure(bg='mint cream')
    wn.minsize(width=500, height=500)
    wn.geometry("700x600")

    Canvas1 = Canvas(wn)
    Canvas1.config(bg='LightBlue1')
    Canvas1.pack(expand=True, fill=BOTH)
    
    headingFrame1 = Frame(wn, bg='LightBlue1', bd=5)
    headingFrame1.place(relx=0.25, rely=0.1, relwidth=0.5, relheight=0.13)
    headingLabel = Label(headingFrame1, text="Add a Product", fg='grey19', font=('Courier', 15, 'bold'))
    headingLabel.place(relx=0, rely=0, relwidth=1, relheight=1)

    labelFrame = Frame(wn)
    labelFrame.place(relx=0.1, rely=0.4, relwidth=0.8, relheight=0.4)
        
    lable1 = Label(labelFrame, text="Date : ", fg='black')
    lable1.place(relx=0.05, rely=0.3, relheight=0.08)
        
    date = Entry(labelFrame)
    date.place(relx=0.3, rely=0.3, relwidth=0.62, relheight=0.08)
        
    lable2 = Label(labelFrame, text="Product Name : ", fg='black')
    lable2.place(relx=0.05, rely=0.45, relheight=0.08)
        
    prodName = Entry(labelFrame)
    prodName.place(relx=0.3, rely=0.45, relwidth=0.62, relheight=0.08)
        
    lable3 = Label(labelFrame, text="Product Price : ", fg='black')
    lable3.place(relx=0.05, rely=0.6, relheight=0.08)
        
    prodPrice = Entry(labelFrame)
    prodPrice.place(relx=0.3, rely=0.6, relwidth=0.62, relheight=0.08)
           
    Btn = Button(wn, text="ADD", bg='#d1ccc0', fg='black', command=prodtoTable)
    Btn.place(relx=0.28, rely=0.85, relwidth=0.18, relheight=0.08)
    
    Quit = Button(wn, text="Quit", bg='#f7f1e3', fg='black', command=wn.destroy)
    Quit.place(relx=0.53, rely=0.85, relwidth=0.18, relheight=0.08)
    
    wn.mainloop()

# Function to remove the product from the Excel file
def removeProd():
    name = prodName.get().lower()
    
    try:
        df = pd.read_excel(products_file)
        df['prodName'] = df['prodName'].str.lower()
        df = df[df['prodName'] != name]
        df.to_excel(products_file, index=False)
        messagebox.showinfo('Success', "Product Record Deleted Successfully")
    except Exception as e:
        print("The exception is:", e)
        messagebox.showinfo("Error", "Please check Product Name")
    
    wn.destroy()

# Function to get product details from the user to be deleted
def delProd():
    global prodName, Canvas1, wn
    
    wn = tkinter.Tk()
    wn.title("PythonGeeks Shop Management System")
    wn.configure(bg='mint cream')
    wn.minsize(width=500, height=500)
    wn.geometry("700x600")

    Canvas1 = Canvas(wn)
    Canvas1.config(bg="misty rose")
    Canvas1.pack(expand=True, fill=BOTH)
    
    headingFrame1 = Frame(wn, bg="misty rose", bd=5)
    headingFrame1.place(relx=0.25, rely=0.1, relwidth=0.5, relheight=0.13)
    headingLabel = Label(headingFrame1, text="Delete Product", fg='grey19', font=('Courier', 15, 'bold'))
    headingLabel.place(relx=0, rely=0, relwidth=1, relheight=1)
    
    labelFrame = Frame(wn)
    labelFrame.place(relx=0.1, rely=0.3, relwidth=0.8, relheight=0.5)   
    
    lable = Label(labelFrame, text="Product Name : ", fg='black')
    lable.place(relx=0.05, rely=0.5)
        
    prodName = Entry(labelFrame)
    prodName.place(relx=0.3, rely=0.5, relwidth=0.62)
    
    Btn = Button(wn, text="DELETE", bg='#d1ccc0', fg='black', command=removeProd)
    Btn.place(relx=0.28, rely=0.9, relwidth=0.18, relheight=0.08)
    
    Quit = Button(wn, text="Quit", bg='#f7f1e3', fg='black', command=wn.destroy)
    Quit.place(relx=0.53, rely=0.9, relwidth=0.18, relheight=0.08)
    
    wn.mainloop()

# Function to show all the products in the Excel file
def viewProds():
    global wn
    
    wn = tkinter.Tk()
    wn.title("PythonGeeks Shop Management System")
    wn.configure(bg='mint cream')
    wn.minsize(width=500, height=500)
    wn.geometry("700x600")

    Canvas1 = Canvas(wn)
    Canvas1.config(bg="old lace")
    Canvas1.pack(expand=True, fill=BOTH)

    headingFrame1 = Frame(wn, bg='old lace', bd=5)
    headingFrame1.place(relx=0.25, rely=0.1, relwidth=0.5, relheight=0.13)

    headingLabel = Label(headingFrame1, text="View Products", fg='black', font=('Courier', 15, 'bold'))
    headingLabel.place(relx=0, rely=0, relwidth=1, relheight=1)
    
    labelFrame = Frame(wn)
    labelFrame.place(relx=0.1, rely=0.3, relwidth=0.8, relheight=0.5)
    y = 0.25

    try:
        df = pd.read_excel(products_file)
        Label(labelFrame, text="%-50s%-50s%-50s"%('Date', 'Product', 'Price'), font=('calibri', 11, 'bold'), fg='black').place(relx=0.07, rely=0.1)
        Label(labelFrame, text="----------------------------------------------------------------------------", fg='black').place (relx=0.05, rely=0.2)
        
        for index, row in df.iterrows():
            Label(labelFrame, text="%-50s%-50s%-50s"%(row["date"], row["prodName"], row["prodPrice"]), fg='black').place(relx=0.07, rely=y)
            y += 0.1
    except Exception as e:
        print("The exception is:", e)
        messagebox.showinfo("Error", "Failed to fetch data from Excel file")
    
    Quit = Button(wn, text="Quit", bg='#f7f1e3', fg='black', command=wn.destroy)
    Quit.place(relx=0.4, rely=0.9, relwidth=0.18, relheight=0.08)
    
    wn.mainloop()

# Function to register sales
def regSale():
    cname = custName.get()
    pname = prodName.get()
    quty = qty.get()
    dt = date.get()

    try:
        df_products = pd.read_excel(products_file)
        df_products['prodName'] = df_products['prodName'].str.lower()
        product = df_products[df_products['prodName'] == pname.lower()]
        price = float(product['prodPrice']) * int(quty)

        new_sale = pd.DataFrame({"custName": [cname], "date": [dt], "prodName": [pname], "qty": [quty], "price": [price]})
        
        df_sales = pd.read_excel(sales_file)
        df_sales = df_sales.append(new_sale, ignore_index=True)
        df_sales.to_excel(sales_file, index=False)
        
        messagebox.showinfo('Success', "Sales Record Added Successfully")
    except Exception as e:
        print("The exception is:", e)
        messagebox.showinfo("Error", "Please check Product Name")
    
    wn.destroy()

# Function to get sales details from the user
def saleProd():
    global prodName, date, qty, custName, Canvas1, wn
    
    wn = tkinter.Tk()
    wn.title("PythonGeeks Shop Management System")
    wn.configure(bg='mint cream')
    wn.minsize(width=500, height=500)
    wn.geometry("700x600")

    Canvas1 = Canvas(wn)
    Canvas1.config(bg="misty rose")
    Canvas1.pack(expand=True, fill=BOTH)

    headingFrame1 = Frame(wn, bg="misty rose", bd=5)
    headingFrame1.place(relx=0.25, rely=0.1, relwidth=0.5, relheight=0.13)
    headingLabel = Label(headingFrame1, text="Register a Sale", fg='black', font=('Courier', 15, 'bold'))
    headingLabel.place(relx=0, rely=0, relwidth=1, relheight=1)

    labelFrame = Frame(wn, bg='misty rose')
    labelFrame.place(relx=0.1, rely=0.3, relwidth=0.8, relheight=0.5)
        
    lable1 = Label(labelFrame, text="Customer Name : ", fg='black', bg='misty rose')
    lable1.place(relx=0.05, rely=0.1, relheight=0.08)
        
    custName = Entry(labelFrame)
    custName.place(relx=0.3, rely=0.1, relwidth=0.62, relheight=0.08)
        
    lable2 = Label(labelFrame, text="Date : ", fg='black', bg='misty rose')
    lable2.place(relx=0.05, rely=0.25, relheight=0.08)
        
    date = Entry(labelFrame)
    date.place(relx=0.3, rely=0.25, relwidth=0.62, relheight=0.08)
        
    lable3 = Label(labelFrame, text="Product Name : ", fg='black', bg='misty rose')
    lable3.place(relx=0.05, rely=0.4, relheight=0.08)
        
    prodName = Entry(labelFrame)
    prodName.place(relx=0.3, rely=0.4, relwidth=0.62, relheight=0.08)
        
    lable4 = Label(labelFrame, text="Quantity : ", fg='black', bg='misty rose')
    lable4.place(relx=0.05, rely=0.55, relheight=0.08)
        
    qty = Entry(labelFrame)
    qty.place(relx=0.3, rely=0.55, relwidth=0.62, relheight=0.08)
    
    Btn = Button(wn, text="ADD", bg='#d1ccc0', fg='black', command=regSale)
    Btn.place(relx=0.28, rely=0.85, relwidth=0.18, relheight=0.08)
    
    Quit = Button(wn, text="Quit", bg='#f7f1e3', fg='black', command=wn.destroy)
    Quit.place(relx=0.53, rely=0.85, relwidth=0.18, relheight=0.08)
    
    wn.mainloop()

# Main Menu
root = tkinter.Tk()
root.title("PythonGeeks Shop Management System")
root.configure(bg='mint cream')
root.minsize(width=400, height=400)
root.geometry("600x500")

headingFrame1 = Frame(root, bg='PaleGreen1', bd=5)
headingFrame1.place(relx=0.25, rely=0.1, relwidth=0.5, relheight=0.13)

headingLabel = Label(headingFrame1, text="Welcome to PythonGeeks", fg='green', font=('Courier', 15, 'bold'))
headingLabel.place(relx=0, rely=0, relwidth=1, relheight=1)

btn1 = Button(root, text="Add a Product", bg='PaleGreen1', fg='green', command=addProd)
btn1.place(relx=0.28, rely=0.3, relwidth=0.45, relheight=0.1)

btn2 = Button(root, text="Delete a Product", bg='PaleGreen1', fg='green', command=delProd)
btn2.place(relx=0.28, rely=0.4, relwidth=0.45, relheight=0.1)

btn3 = Button(root, text="View Products", bg='PaleGreen1', fg='green', command=viewProds)
btn3.place(relx=0.28, rely=0.5, relwidth=0.45, relheight=0.1)

btn4 = Button(root, text="Register a Sale", bg='PaleGreen1', fg='green', command=saleProd)
btn4.place(relx=0.28, rely=0.6, relwidth=0.45, relheight=0.1)

root.mainloop()
