import tkinter as tk
from tkinter import ttk
from docxtpl import DocxTemplate
import datetime
from tkinter import messagebox
from tkinter import PhotoImage
import win32api
import sqlite3
from ttkwidgets.autocomplete import AutocompleteEntry
import os
import num2words

invoicenumber=1
invoice_list = []
global items
items=[]
printpath=''

PRIMARY_COLOR='#003083'
SECONDARY_COLOR='#281E5D'
PRIMARY_ACCENT='#87CEEB'
SECONDARY_ACCENT='#000000'

def initialize_item():
    try:
        #conn = sqlite3.connect('firsttable.db')
        conn = sqlite3.connect('secondtable.db')
        cursor = conn.cursor()
        items.clear()  
        query = ("SELECT * FROM ITEMS")
        cursor.execute(query)
            
        # Fetch and output result
        results = cursor.fetchall()
        for result in results:
            items.append(result[0])
        cursor.close()
        
        if not os.path.exists("C:/Bills"): 
            os.makedirs("C:/Bills")
    except sqlite3.Error as error:
        print('Error occurred - ',error)


initialize_item()

class MultiFrameApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Invoicer")
        self.geometry("1024x654")
        self.frames = {}
        self.resizable(0, 0)
        photo = PhotoImage(file = "info.png")
        self.iconphoto(False, photo)
        self.create_sidebar()
        self.create_frames()
        self.show_frame("Frame1")
        self.update_date()  # Initial update
        self.schedule_date_update()
        self.view_item()
        self.connect()
        self.qty_spinbox.delete(0, tk.END)
        self.qty_spinbox.insert(0, "1")
    
    def create_sidebar(self):
        sidebar = tk.Frame(self, bg="lightgray", width=200)
        sidebar.grid(row=0, column=0, sticky="ns")

        button_bg = SECONDARY_COLOR
        button_fg = "#F4F3EE"
        button_font = ("Rockwell", 16)
        x = 4
        y = 12

        button1 = tk.Button(sidebar, text="Create Invoice", command=lambda: self.show_frame("Frame1"), height=x, width=y, font=button_font, bg=button_bg, fg=button_fg)
        button2 = tk.Button(sidebar, text="View Items", command=lambda: self.show_frame("Frame2"), height=x, width=y, font=button_font, bg=button_bg, fg=button_fg)
        button3 = tk.Button(sidebar, text="+ / - Item", command=lambda: self.show_frame("Frame3"), height=x, width=y, font=button_font, bg=button_bg, fg=button_fg)
        button4 = tk.Button(sidebar, text="Update Item", command=lambda: self.show_frame("Frame4"), height=x, width=y, font=button_font, bg=button_bg, fg=button_fg)
        button5 = tk.Button(sidebar, text="View Invoice", command=lambda: self.show_frame("Frame5"), height=x, width=y, font=button_font, bg=button_bg, fg=button_fg)
        #button6 = tk.Button(sidebar, text="Sales Data", command=lambda: self.show_frame("Frame6"), height=x, width=y, font=button_font, bg=button_bg, fg=button_fg)
        button6 = tk.Button(sidebar, text="Sales Data", command=lambda: (self.show_frame("Frame6"), self.Sales()), height=x, width=y, font=button_font, bg=button_bg, fg=button_fg)


        button1.grid(row=0, sticky="ew")
        button2.grid(row=1, sticky="ew")
        button3.grid(row=2, sticky="ew")
        button4.grid(row=3, sticky="ew")
        button5.grid(row=4, sticky="ew")
        button6.grid(row=5, sticky="ew")

        sidebar.grid_rowconfigure(0, weight=1)

    def create_frames(self):
        frame1 = tk.Frame(self, bg=PRIMARY_COLOR)

        font_size=12
        invoice_label = tk.Label(frame1,font=('Rockwell Bold', font_size), text="Invoice no.",bg=PRIMARY_COLOR)
        invoice_label.grid(row=0, column=0,sticky="w",padx=20,pady=10)
        self.invoice_entry = tk.Entry(frame1,font=('Rockwell Bold', font_size),width=30)
        self.invoice_entry.grid(row=1, column=0,sticky="w",padx=20,pady=3)
        invoicenumstring = self.get_invoice_number()
        self.invoice_entry.insert(0,invoicenumstring)

        
        first_name_label = tk.Label(frame1,font=('Rockwell Bold', font_size), text="Name of Customer",bg=PRIMARY_COLOR)
        first_name_label.grid(row=2, column=0,sticky="w",padx=20,pady=3)
        self.first_name_entry = tk.Entry(frame1,font=('Rockwell Bold', font_size),width=30)
        self.first_name_entry.grid(row=3, column=0,sticky="w",padx=20,pady=3)


        phone_label = tk.Label(frame1,font=('Rockwell Bold', font_size), text="Phone",bg=PRIMARY_COLOR)
        phone_label.grid(row=4, column=0,sticky="w",padx=20,pady=3)
        self.phone_entry = tk.Entry(frame1,font=('Rockwell Bold', font_size),width=30)
        self.phone_entry.grid(row=5, column=0,sticky="w",padx=20,pady=3)

        self.desc_label = tk.Label(frame1,font=('Rockwell Bold', font_size), text="Description",bg=PRIMARY_COLOR)
        self.desc_label.grid(row=6, column=0,sticky="w",padx=20,pady=3)
        #self.desc_entry = tk.Entry(frame1)

        qty_label = tk.Label(frame1,font=('Rockwell Bold', font_size), text="Qty",bg=PRIMARY_COLOR)
        qty_label.grid(row=8, column=0,sticky="w",padx=20,pady=3)
        self.qty_spinbox = tk.Spinbox(frame1,font=('Rockwell Bold', font_size), from_=0.0, to=100.0,increment=0.1,width = 29)
        self.qty_spinbox.grid(row=9, column=0,sticky="w",padx=(20,0),pady=3)

        #column 2
        date_label = tk.Label(frame1,font=('Rockwell Bold', font_size), text="Date",bg=PRIMARY_COLOR)
        date_label.grid(row=0, column=2,sticky="e",padx=20,pady=3)
        self.date_entry = tk.Entry(frame1,font=('Rockwell Bold', font_size),width=20)
        self.date_entry.grid(row=1, column=2,sticky="e",padx=20,pady=3)
        self.date_entry.insert(0,datetime.datetime.now().strftime("%d/%m/%Y %H:%M"))
        
        discount_label = tk.Label(frame1,font=('Rockwell Bold', font_size), text="Discount",bg=PRIMARY_COLOR)
        discount_label.grid(row=2, column=2,sticky="e",padx=20,pady=3)
        self.discount_entry = tk.Entry(frame1,font=('Rockwell Bold', font_size),width=20)
        self.discount_entry.grid(row=3, column=2,sticky="e",padx=20,pady=3)
        self.discount_entry.insert(0,0)#insert 0 as default value for discount

        tax_label = tk.Label(frame1,font=('Rockwell Bold', font_size), text="Tax",bg=PRIMARY_COLOR)
        tax_label.grid(row=4, column=2,sticky="e",padx=20,pady=3)
        self.tax_entry = tk.Entry(frame1,font=('Rockwell Bold', font_size),width=20)
        self.tax_entry.grid(row=5, column=2,sticky="e",padx=20,pady=3)
        self.tax_entry.insert(0,0)#insert 0 as default value for tax

        self.desc_entry = AutocompleteEntry(frame1,font=('Rockwell Bold', font_size),completevalues=items,width=30)
        self.desc_entry.grid(row=7,columnspan=3,sticky="w",padx=20,pady=3)
        self.desc_entry.bind('<Tab>', self.set_price)


        name_var = tk.StringVar()
        self.price_label = tk.Label(frame1,font=('Rockwell Bold', font_size), text="Unit Price",bg=PRIMARY_COLOR)
        self.price_label.grid(row=6, column=2,sticky="e",padx=20,pady=3)
        self.price_spinbox  = tk.Entry(frame1,font=('Rockwell Bold', font_size), textvariable=name_var)
        self.price_spinbox.grid(row=7, column=2,sticky="e",padx=20,pady=3)
        #self.price_spinbox = tk.Spinbox(frame1, from_=0.0, to=500, increment=0.5)
        #self.price_spinbox.grid(row=3, column=2)

        button_bg = SECONDARY_ACCENT
        button_fg = "#F4F3EE"
        button_font = ("Rockwell",9)
        button_width = 25

        add_item_button = tk.Button(frame1, text = "Add Item", command = self.add_item,width=button_width,font=button_font, bg=button_bg, fg=button_fg)
        add_item_button.grid(row=9, column=1,padx=(0,10), pady=5)

        remove_item_button = tk.Button(frame1, text="  Remove Item  ", command=self.remove_item,width=button_width,font=button_font, bg=button_bg, fg=button_fg)
        remove_item_button.grid(row=9, column=2,sticky="e",pady=3,padx=20)


        #Trying to style treeview now
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview.Heading", background=SECONDARY_COLOR, foreground="white")
        style.configure("Treeview",background=PRIMARY_ACCENT,foreground="black",fieldbackground="white",font=("Rockwell Bold", 10))
        

        
        columns = ('qty', 'desc', 'price', 'total')
        self.tree = ttk.Treeview(frame1, columns=columns, show="headings")
        self.tree.heading('qty', text='Qty')
        self.tree.heading('desc', text='Description')
        self.tree.heading('price', text='Unit Price')
        self.tree.heading('total', text="Total")
        self.tree.tag_configure('oddrow', background=PRIMARY_ACCENT)
        self.tree.tag_configure('evenrow', background='white')

            
        self.tree.grid(row=10, column=0, columnspan=3, padx=20, pady=10)    


        self.total_amount = tk.StringVar()
        self.total_amount.set("0.00")  # Initialize it with 0.00

        # Create an Entry widget to display the total amount
        total_label = tk.Label(frame1, text="Total Amount:", font=('Rockwell Bold', 12), bg=PRIMARY_COLOR)
        total_label.grid(row=11, column=1, sticky="w", padx=20, pady=3)
        self.total_entry = tk.Entry(frame1, font=('Rockwell Bold', 12), textvariable=self.total_amount, width=20)
        self.total_entry.grid(row=11, column=2, columnspan=1, sticky="e", padx=20, pady=3)

        new_invoice_button = tk.Button(frame1, text="New Invoice", command=self.new_invoice,width=button_width ,font=button_font, bg=button_bg, fg=button_fg)
        new_invoice_button.grid(row=12, column=0, sticky="w",padx=(20, 0), pady=5)
        save_invoice_button = tk.Button(frame1, text="Generate Invoice", command=self.generate_invoice,width=button_width, font=button_font, bg=button_bg, fg=button_fg)
        save_invoice_button.grid(row=12, column=1, padx=(0,10), pady=5)  
        print_invoice_button = tk.Button(frame1, text="Print Invoice", command=self.print_invoice,width=button_width,font=button_font, bg=button_bg, fg=button_fg)
        print_invoice_button.grid(row=12, column=2, sticky="e", padx=(0, 20), pady=5)


        self.frames["Frame1"] = frame1


        # Define other frames (Frame2 to Frame6) similarly
        # FRAME 2
        frame2 = tk.Frame(self, bg=PRIMARY_COLOR)
    
        label2 = tk.Label(frame2, text="View All Items", font=("Rockwell Bold", 20), bg=PRIMARY_COLOR)
        label2.grid(row=0, column=0, pady=10, columnspan=3)  # Use grid instead of pack

        # Define the content for Frame2 heres

        columns = ('desc', 'price', 'qty')
        self.tree1 = ttk.Treeview(frame2, columns=columns, show="headings",height=25)
        self.tree1.heading('desc', text='Description')
        self.tree1.heading('price', text='Unit Price')
        self.tree1.heading('qty', text='Qty')

        self.tree1.column("desc", width=260)
        self.tree1.column("price", width=260)
        self.tree1.column("qty", width=260)
        self.tree1.grid(row=2, column=0, columnspan=3, padx=(20, 0), pady=10, sticky="ew")
        self.tree1.tag_configure('oddrow', background=PRIMARY_ACCENT)
        self.tree1.tag_configure('evenrow', background='white')

        scrollbar = ttk.Scrollbar(frame2, orient="vertical", command=self.tree1.yview)
        # Place the scrollbar on the right side of the Treeview
        scrollbar.grid(row=2, column=3, sticky="ns",pady=10)
        # Connect scrollbar to Treeview's yview method
        self.tree1.configure(yscrollcommand=scrollbar.set)
        self.frames["Frame2"] = frame2


        #FRAME 3
        frame3 = tk.Frame(self, bg=PRIMARY_COLOR)
        label3 = tk.Label(frame3, text="Add Item", font=("Rockwell Bold", 20), bg=PRIMARY_COLOR)
        label3.pack(pady=10)

        # Define the content for Frame3 here
        label3a= tk.Label(frame3, text="Input Item Name", font=("Rockwell", 14),bg=PRIMARY_COLOR)
        label3a.pack(padx=20, pady=5, anchor="w")
        self.entry3a = tk.Entry(frame3, font=("Rockwell", 14),width=30)
        self.entry3a.pack(padx=20, pady=4, anchor="w")

        label3b = tk.Label(frame3, text="Input Item Price", font=("Rockwell", 14),bg=PRIMARY_COLOR)
        label3b.pack(padx=20, pady=5, anchor="w")
        self.entry3b = tk.Entry(frame3, font=("Rockwell", 14),width=30)
        self.entry3b.pack(padx=20, pady=4, anchor="w")

        label3c = tk.Label(frame3, text="Input Item Quantity", font=("Rockwell", 14),bg=PRIMARY_COLOR)
        label3c.pack(padx=20, pady=4, anchor="w")
        self.entry3c = tk.Entry(frame3, font=("Rockwell", 14),width=30)
        self.entry3c.pack(padx=20, pady=4, anchor="w")

        button31 = tk.Button(frame3, text="Insert", font=("Rockwell", 14),command=self.add,bg=button_bg, fg=button_fg)
        button31.pack(padx=20, pady=8, anchor="w")


        ################
        label3d = tk.Label(frame3, text="Remove Item", font=("Rockwell Bold", 20), bg=PRIMARY_COLOR)
        label3d.pack(pady=10)

        # Define the content for Frame5 here
        label3e= tk.Label(frame3, text="Input Item Name to Remove", font=("Rockwell", 14),bg=PRIMARY_COLOR)
        label3e.pack(padx=20, pady=4, anchor="w")
        #entry5a = tk.Entry(frame5, font=("Rockwell", 14))
        self.entry3f = AutocompleteEntry(frame3,font=('Rockwell Bold', font_size),completevalues=items,width=33)
        self.entry3f.pack(padx=20, pady=4, anchor="w")
        self.entry3f.bind('<Return>', self.set_price2)

        label3g = tk.Label(frame3, text="Item Price", font=("Rockwell", 14),bg=PRIMARY_COLOR)
        label3g.pack(padx=20, pady=4, anchor="w")
        self.entry3h = tk.Entry(frame3, font=("Rockwell", 14),width=30)
        self.entry3h.pack(padx=20, pady=4, anchor="w")

        label3i = tk.Label(frame3, text="Number of Items to be Removed", font=("Rockwell", 14),bg=PRIMARY_COLOR)
        label3i.pack(padx=20, pady=4, anchor="w")
        self.entry3j = tk.Entry(frame3, font=("Rockwell", 14),width=30)
        self.entry3j.pack(padx=20, pady=4, anchor="w")

        button32 = tk.Button(frame3, text="Remove Item", font=("Rockwell", 14),command=self.delete, bg=button_bg, fg=button_fg)
        button32.pack(padx=20, pady=8, anchor="w")
        ################


        self.frames["Frame3"] = frame3


        #FRAME 4
        frame4 = tk.Frame(self, bg=PRIMARY_COLOR)
        label4 = tk.Label(frame4, text="Update Item", font=("Rockwell Bold", 20), bg=PRIMARY_COLOR)
        label4.pack(pady=10)

        # Define the content for Frame4 here
        label4a= tk.Label(frame4, text="Input Item Name", font=("Rockwell", 14),bg=PRIMARY_COLOR)
        label4a.pack(padx=20, pady=5, anchor="w")
        #entry4a = tk.Entry(frame4, font=("Rockwell", 14))
        self.entry4a = AutocompleteEntry(frame4,font=('Rockwell Bold', font_size),completevalues=items,width=33)
        self.entry4a.pack(padx=20, pady=5, anchor="w")
        self.entry4a.bind('<Return>', self.set_price1)


        label4b = tk.Label(frame4, text="Old Price for Item", font=("Rockwell", 14),bg=PRIMARY_COLOR)
        label4b.pack(padx=20, pady=5, anchor="w")
        self.entry4b = tk.Entry(frame4, font=("Rockwell", 14),width=30)
        self.entry4b.pack(padx=20, pady=5, anchor="w")
        
        label4c = tk.Label(frame4, text="Input New Price for Item", font=("Rockwell", 14),bg=PRIMARY_COLOR)
        label4c.pack(padx=20, pady=5, anchor="w")
        self.entry4c = tk.Entry(frame4, font=("Rockwell", 14),width=30)
        self.entry4c.pack(padx=20, pady=5, anchor="w")


        button42= tk.Button(frame4, text="Update Price", font=("Rockwell", 14),command=self.update, bg=button_bg, fg=button_fg)
        button42.pack(padx=20, pady=10, anchor="w")

        self.frames["Frame4"] = frame4


        #FRAME 5
        frame5 = tk.Frame(self, bg=PRIMARY_COLOR)
        label5 = tk.Label(frame5, text="View Bill", font=("Rockwell Bold", 20), bg=PRIMARY_COLOR)
        label5.grid(row=0, column=0, pady=10, columnspan=3)
        
        # Define the content for Frame5 here
        label5b = tk.Label(frame5, text="Enter Invoice Number", font=("Rockwell", 14), bg=PRIMARY_COLOR)
        label5b.grid(row=1, column=0, padx=(20,10), pady=5, sticky="w")

        self.entry5b = tk.Entry(frame5, font=("Rockwell", 14))
        self.entry5b.grid(row=1, column=1, padx=(0,0), pady=5, sticky="w")

        button51 = tk.Button(frame5, text="Fetch Bill", font=("Rockwell", 11), command=self.get_invoice, bg=button_bg, fg=button_fg)
        button51.grid(row=1, column=2, padx=(10,0), pady=10, sticky="e")

        # Treeview for 5th Frame
        columns = ('desc', 'price', 'qty')
        self.tree2 = ttk.Treeview(frame5, columns=columns, show="headings",height=23)
        self.tree2.heading('desc', text='Description')
        self.tree2.heading('price', text='Unit Price')
        self.tree2.heading('qty', text='Qty')

        self.tree2.column("desc", width=260)
        self.tree2.column("price", width=260)
        self.tree2.column("qty", width=260)
        self.tree2.grid(row=3, column=0, columnspan=3, padx=(20, 0), pady=10, sticky="ew")
        self.tree2.tag_configure('oddrow', background=PRIMARY_ACCENT)
        self.tree2.tag_configure('evenrow', background='white')

        scrollbar2 = ttk.Scrollbar(frame5, orient="vertical", command=self.tree2.yview)
        # Place the scrollbar on the right side of the Treeview
        scrollbar2.grid(row=3, column=3, sticky="ns",pady=10)
        # Connect scrollbar to Treeview's yview method
        self.tree2.configure(yscrollcommand=scrollbar2.set)

        self.frames["Frame5"] = frame5


        # FRAME 6
        frame6 = tk.Frame(self, bg=PRIMARY_COLOR)
        label6 = tk.Label(frame6, text="Daily Sales", font=("Rockwell Bold", 20), bg=PRIMARY_COLOR)
        label6.pack(pady=10)

        # Define the content for Frame6 here
        label6a= tk.Label(frame6, text="Total Sales Today", font=("Rockwell", 14), bg=PRIMARY_COLOR)
        label6a.pack(padx=20, pady=10, anchor="w")
        #label6b= tk.Label(frame6, text="Value", font=("Rockwell", 14), bg=PRIMARY_COLOR)
        #label6b.pack(padx=20, pady=10, anchor="w")
        self.entry6b= tk.Entry(frame6, font=("Rockwell", 14),width=30)
        self.entry6b.pack(padx=20, pady=10, anchor="w")

        label6c = tk.Label(frame6, text="Number of Items Sold", font=("Rockwell", 14), bg=PRIMARY_COLOR)
        label6c.pack(padx=20, pady=10, anchor="sw")
        self.entry6d= tk.Entry(frame6, font=("Rockwell", 14),width=30)
        self.entry6d.pack(padx=20, pady=10, anchor="w")

        label6e = tk.Label(frame6, text="Highest Invoice Value Today", font=("Rockwell", 14), bg=PRIMARY_COLOR)
        label6e.pack(padx=20, pady=10, anchor="sw")
        self.entry6f= tk.Entry(frame6, font=("Rockwell", 14),width=30)
        self.entry6f.pack(padx=20, pady=10, anchor="w")
        
        label6g = tk.Label(frame6, text="Number of Invoices Generated Today", font=("Rockwell", 14), bg=PRIMARY_COLOR)
        label6g.pack(padx=20, pady=10, anchor="sw")
        self.entry6h= tk.Entry(frame6, font=("Rockwell", 14),width=30)
        self.entry6h.pack(padx=20, pady=10, anchor="w")

        label6i = tk.Label(frame6, text="Application By:", font=("Rockwell", 14), bg=PRIMARY_COLOR)
        label6i.pack(padx=20, pady=(120,10), anchor="e")
        label6j = tk.Label(frame6, text="Akash Aparajeya", font=("Rockwell", 14), bg=PRIMARY_COLOR)
        label6j.pack(padx=20, pady=10, anchor="e")

        self.frames["Frame6"] = frame6
        # Configure column to expand horizontally
        self.grid_columnconfigure(1, weight=1)

    def show_frame(self, frame_name):
        frame = self.frames[frame_name]
        frame.grid(row=0, column=1, sticky="nsew")
        frame.tkraise()
            
    def Sales(self):
        conn = sqlite3.connect('secondtable.db')
        cursor = conn.cursor()
        # Get the current date
        current_datetime = datetime.datetime.now()  # Get current date and time
        next_date = (current_datetime + datetime.timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
        # Set time component to midnight (00:00:00)
        current_date = current_datetime.replace(hour=0, minute=0, second=0, microsecond=0)
        print(current_date)
        # Query to retrieve total sales amount for the current day
        query = "SELECT SUM(Amount) FROM ORDERS WHERE Date >= ? AND Date <=?"
        cursor.execute(query, (current_date,next_date))
        # Fetch the result
        total_sales = cursor.fetchone()[0]
        print("TOTAL SALES IS ",total_sales)

        ###
        query = "SELECT SUM(Quantity) FROM ORDERITEMS WHERE OrderId IN (SELECT InvoiceNum FROM ORDERS WHERE Date >= ? AND Date<=?)"
        cursor.execute(query, (current_date,next_date))
        # Fetch the result
        total_items_sold = cursor.fetchone()[0]
        ###
        query = "SELECT MAX(Amount) FROM ORDERS WHERE Date >= ? AND Date<=?"
        cursor.execute(query, (current_date,next_date))
        # Fetch the result
        highest_invoice_value = cursor.fetchone()[0]

        ###
        query = "SELECT COUNT(DISTINCT InvoiceNum) FROM ORDERS WHERE Date >= ? AND Date <=?"
        cursor.execute(query, (current_date,next_date))
        # Fetch the result
        number_of_invoices = cursor.fetchone()[0]

        self.entry6b.delete(0, tk.END)
        self.entry6d.delete(0, tk.END)
        self.entry6f.delete(0, tk.END)
        self.entry6h.delete(0, tk.END)
        self.entry6b.insert(0,str(total_sales))
        self.entry6d.insert(0,str(total_items_sold))
        self.entry6f.insert(0,str(highest_invoice_value))
        self.entry6h.insert(0,str(number_of_invoices))
        

    def clear_item(self):
        self.qty_spinbox.delete(0, tk.END)
        self.qty_spinbox.insert(0, "1")
        self.desc_entry.delete(0, tk.END)
        self.price_spinbox.delete(0, tk.END)
        self.price_spinbox.insert(0, "0.0")
        self.total_entry.delete(0, tk.END)
        self.total_entry.insert(0, "0.0")

    def add_item(self):
        qty = float(self.qty_spinbox.get())
        desc = self.desc_entry.get().capitalize()
        price = float(self.price_spinbox.get())
        line_total = round((qty*price),2)
        invoice_item = [qty, desc, price, line_total]
        #self.tree.insert('','end', values=invoice_item)
        tag = 'evenrow' if len(self.tree.get_children()) % 2 == 0 else 'oddrow'
        self.tree.insert('', 'end', values=invoice_item, tags=(tag,))
        
        self.clear_item()
        invoice_list.append(invoice_item)
        self.calculate_total()

    def calculate_total(self):
        discc = int(self.discount_entry.get())
        total = sum(float(self.tree.item(item, "values")[3]) for item in self.tree.get_children())
        total = round(((total*(100-discc))/100),2)
        self.total_amount.set(f"{total:.2f}")
        
    def new_invoice(self):
        self.first_name_entry.delete(0, tk.END)
        self.phone_entry.delete(0, tk.END)
        self.clear_item()
        self.tree.delete(*self.tree.get_children())
        
        invoice_list.clear()

    def get_invoice(self):
        invoiceId = str(self.entry5b.get())
        for item in self.tree2.get_children():
            self.tree2.delete(item)
        conn = sqlite3.connect('secondtable.db')
        cursor = conn.cursor()
        query = "SELECT * FROM ORDERS WHERE InvoiceNum=?"
        cursor.execute(query,(invoiceId,))
        #cursor.execute(query)
        result = cursor.fetchall()
        if result:
            for res in result:
                print(res)
                CustomerName=res[2]
                Total=res[3]
                Date=res[4]

        else:
            print("Nothing is here son")
            msg1="Unable to find the Invoice with that InvoiceId"
            messagebox.showinfo("No Invoice Found", msg1)
            return
        CustomerNameString = ("Customer Name "+CustomerName)
        TotalString = ("Total Amount "+str(Total))
        DateString = ("DateTime "+str(Date))
        tag = 'evenrow' if len(self.tree1.get_children()) % 2 == 0 else 'oddrow'
        self.tree2.insert('', 'end', values=CustomerNameString, tags=(tag,))
        
        query = "SELECT * FROM ORDERITEMS WHERE OrderId=?"
        results = cursor.execute(query,(invoiceId,))
        for result in results:
            bill_item=[]
            bill_item.append(result[1])
            bill_item.append(result[2])
            bill_item.append(result[3])
            tag = 'evenrow' if len(self.tree2.get_children()) % 2 == 0 else 'oddrow'
            self.tree2.insert('', 'end', values=bill_item, tags=(tag,))
        tag = 'evenrow' if len(self.tree2.get_children()) % 2 == 0 else 'oddrow'
        self.tree2.insert('', 'end', values=TotalString, tags=(tag,))
        tag = 'evenrow' if len(self.tree2.get_children()) % 2 == 0 else 'oddrow'
        self.tree2.insert('', 'end', values=DateString, tags=(tag,))

        
        cursor.close()    
        
        print("sure")
        
    def generate_invoice(self):
        global invoicenumber,printpath
        doc = DocxTemplate("invoice_template.docx")
        name = self.first_name_entry.get()
        phone = self.phone_entry.get()
        subtotal = sum(item[3] for item in invoice_list) 

        tax = int(self.tax_entry.get())
        discount = int(self.discount_entry.get())
        disprice = ((100-discount)*subtotal)/100
        disprice = round(disprice,2)
        taxamt = ((tax*disprice)/(100+tax))
        taxamt = round(taxamt,2)
        price = ((100*disprice)/(100+tax))
        price = round(price,2)
        total = taxamt+price
        total = round(total,2)

        datee = self.date_entry.get()
        invoicenum=str(self.invoice_entry.get())

        retval = self.update_order_db(name,invoicenum,total)
        if(retval==1):
            msg1="Unable to Create Invoice Due to Insufficient Quantity"
            messagebox.showinfo("Invoice Failed", msg1)
            return
        
        srno=1
        for item in invoice_list:
            item.append(srno)
            srno=srno+1

        numword = num2words.num2words(total, lang='en_IN')
        numword = numword.capitalize()
        numword = str(numword)+" rupees only"
        doc.render({
                "datee":datee,
                "name":name,
                "phone":phone,
                "invnum":invoicenum,
                "invoice_list": invoice_list,
                "subtotal":subtotal,
                "amt":str(discount)+"%",
                "afterdis":disprice,
                "price":price,
                "Total":total,
                "Numword":numword
                })
        
        cur_date = datetime.datetime.now()
        res = cur_date.strftime("%B")
        resstr= str(res)
        path = "C:/Bills/"+resstr+"/"
        
        if not os.path.exists(path):  # Create the directory if it doesn't exist
            os.mkdir(path)
            file_path = "integer.txt"
            
            with open(file_path, "r") as file:
                invoicenumber = int(file.read())
            
            invoicenumber = 1 #reset invoice number to 1 as a new month has arrived

            with open(file_path, "w") as file:
                file.write(str(invoicenumber))
                
        invoicenumber=str(self.invoice_entry.get())
        
        doc_name=invoicenum+" "+datetime.datetime.now().strftime("%Y-%m-%d") +".docx"
        temp=path+doc_name
        printpath=temp
        doc.save(temp)

        msg1="Invoice Saved in C:/Bills"
        messagebox.showinfo("Invoice Complete", msg1)
        self.view_item()
        #self.increment_invoicenum()
        invoicenumstring = self.generate_invoice_number()
        #invoicenumstring=str(invoicenumber)+datetime.datetime.now().strftime("D%d")
        self.invoice_entry.delete(0,'end')
        #print(invoicenumstring)
        self.invoice_entry.insert(0,invoicenumstring)
        self.new_invoice()

    def print_invoice(self):
        self.generate_invoice()
        #file_to_print ="C:/E 118GB/Tkinter/Office 2.0/inv.docx"
        file_to_print = printpath
        if file_to_print:
            win32api.ShellExecute(0, "print", file_to_print, None, ".", 0)


    def connect(self):
        try:
            conn = sqlite3.connect('secondtable.db')
            cursor = conn.cursor()
            print("hoho")
            #cursor.execute("DROP TABLE ORDERITEMS;")
            # Check if ITEMS table already exists
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='ITEMS';")
            if cursor.fetchone():
                print("ITEMS table already exists.")
            else:
                query = """ CREATE TABLE ITEMS (
                    Item VARCHAR(355) NOT NULL,
                    Price float NOT NULL,
                    Quantity float NOT NULL
                ); """
                cursor.execute(query)
                conn.commit()
                print("ITEMS table created.")

            # Check if ORDERS table already exists
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='ORDERS';")
            if cursor.fetchone():
                print("ORDERS table already exists.")
            else:
                query = """ CREATE TABLE ORDERS (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    InvoiceNum VARCHAR(355) NOT NULL,
                    Name VARCHAR(255) NOT NULL,
                    Amount float NOT NULL,
                    Date time NOT NULL
                ); """
                cursor.execute(query)
                conn.commit()
                print("ORDERS table created.")

            # Check if ORDERITEMS table already exists
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='ORDERITEMS';")
            if cursor.fetchone():
                print("ORDERITEMS table already exists.")
            else:
                query = """ CREATE TABLE ORDERITEMS (
                    OrderId VARCHAR(355) NOT NULL,
                    ItemName VARCHAR(255) NOT NULL,
                    Price float NOT NULL,
                    Quantity float NOT NULL,
                    FOREIGN KEY (OrderId) REFERENCES ORDERS(Id)
                ); """
                cursor.execute(query)
                conn.commit()
                print("ORDERITEMS table created.")


            #######
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='SequenceNumber';")
            if cursor.fetchone():
                print("SequenceNumber table already exists.")
            else:
                query = """ CREATE TABLE SequenceNumber (
                                id INTEGER PRIMARY KEY,
                                last_date DATE,
                                number INTEGER
                ); """
                cursor.execute(query)
                conn.commit()
                print("SequenceNumber table created.")
            ######
                
            cursor.close()
            conn.commit()
            conn.close()
        except sqlite3.Error as error:
            print('Error occurred - ', error)

    
    def find(self):
        try:    
            #conn = sqlite3.connect('firsttable.db')
            conn = sqlite3.connect('secondtable.db')
            cursor = conn.cursor()

            IName = input()
            mylist=[]
            mylist.append(IName)
            MyTuple=tuple(mylist)
            mylist.clear()
            
            query = ("SELECT * FROM ITEMS WHERE Item = ?")
            cursor.execute(query,MyTuple)

            result = cursor.fetchone()

            cursor.close()
        except sqlite3.Error as error:
            print('Error occurred - ',error)

    def view_item(self):
        try:
            for item in self.tree1.get_children():
                self.tree1.delete(item)
                
            templist=[]
            #conn = sqlite3.connect('firsttable.db')
            conn = sqlite3.connect('secondtable.db')
            cursor = conn.cursor()

            query = ("SELECT * FROM ITEMS")
            cursor.execute(query)

            # Fetch and output result
            results = cursor.fetchall()
            for result in results:
                desc=result[0]
                price=result[1]
                qty=result[2]
                templist=[desc,price,qty]
                #self.tree1.insert('','end', values=templist)
                tag = 'evenrow' if len(self.tree1.get_children()) % 2 == 0 else 'oddrow'
                self.tree1.insert('', 'end', values=templist, tags=(tag,))
        
            cursor.close()
        except sqlite3.Error as error:
            print('Error occurred - ',error)

    def update_order_db(self, name, invoicenum, total):
        conn = sqlite3.connect('secondtable.db')
        cursor = conn.cursor()

        # Insert into ORDERS table
        current_timestamp = datetime.datetime.now()
        query = "INSERT INTO ORDERS (InvoiceNum, Name, Amount, Date) VALUES (?, ?, ?, ?)"
        cursor.execute(query, (invoicenum, name, total, current_timestamp))
        conn.commit()

        # Insert into ORDERITEMS table
        for item in invoice_list:
            query = "INSERT INTO ORDERITEMS (OrderId, ItemName, Price, Quantity) VALUES (?, ?, ?, ?)"
            cursor.execute(query, (invoicenum, item[1], item[3], item[0]))
        
        # Update quantities in ITEMS table
        for item in invoice_list:
            # Retrieve current quantity for the item
            query = "SELECT Quantity FROM ITEMS WHERE Item=?"
            cursor.execute(query, (item[1].lower(),))
            row = cursor.fetchone()
            if row:
                current_quantity = row[0]
                sold_quantity = item[0]
                new_quantity = current_quantity - sold_quantity
                if new_quantity < 0:
                    # Optional: Handle cases where sold quantity exceeds available quantity
                    msg1=f"Error: Not enough quantity available for {item[1]}"
                    messagebox.showinfo("Invoice Failed", msg1)
                    return 1
                else:
                    # Update quantity in ITEMS table
                    query = "UPDATE ITEMS SET Quantity=? WHERE Item=?"
                    cursor.execute(query, (new_quantity, item[1].lower()))
                    conn.commit()

        cursor.close()
        conn.close()
        return 0

        
    def add(self):
        try:    
            #conn = sqlite3.connect('firsttable.db')
            conn = sqlite3.connect('secondtable.db')
            cursor = conn.cursor()

            IName = (self.entry3a.get()).lower()
            IPrices = self.entry3b.get()
            IPrice = float(IPrices)
            IQuantity = float(self.entry3c.get())

            mylist=[]
            mylist.append(IName)
            mylist.append(IPrice)
            mylist.append(IQuantity)
            MyTuple=tuple(mylist)

            mytemplist=[]
            mytemplist.append(IName)
            query = ("SELECT COUNT(Item) FROM ITEMS WHERE Item=(?)")
            cursor.execute(query,mytemplist)
            results = cursor.fetchall()
            for result in results:
                if result[0] == 0:
                    query = ("INSERT INTO ITEMS VALUES (?,?,?)")
                    cursor.execute(query,MyTuple)
                    conn.commit()
                    #self.showall()
                    IPrice=str(IPrice)
                    msg1=IName+" with Price "+IPrice+" added to database"
                    messagebox.showinfo("Item Added", msg1)
                 
                    # Close the cursor
                    cursor.close()
                else:
                    msg1="Item with "+IName+" already exists in the database with price "+str(IPrice)
                    messagebox.showinfo("Item Present Already", msg1)
        
            mytemplist.clear()
            self.entry3a.delete(0, tk.END)
            self.entry3b.delete(0, tk.END)
            initialize_item()
            self.view_item()
            #self.create_frames() #TESTING THIS LINE TO FIX ANOMALOUS BEHAVIOUR OF AUTOCOMPLETE
        
        except sqlite3.Error as error:
            print('Error occurred - ',error)

    def showall(self):
        try:    
            #conn = sqlite3.connect('firsttable.db')
            conn = sqlite3.connect('secondtable.db')
            cursor = conn.cursor()            
            query = ("SELECT * FROM ITEMS")
            cursor.execute(query)

            # Fetch and output result
            results = cursor.fetchall()
            for result in results:
                print("Item Name is -" + result[0]+" Item Price is "+str(result[1]))
            # Close the cursor
            cursor.close()
        except sqlite3.Error as error:
            print('Error occurred - ',error)

    
    def delete(self):
        try:    
            #conn = sqlite3.connect('firsttable.db')
            conn = sqlite3.connect('secondtable.db')
            cursor = conn.cursor()

            IName  = self.entry3f.get()
            IQuantity = self.entry3j.get()
            mylist=[]
            mylist.append(IName)
            MyTuple=tuple(mylist)
            mylist.clear()

            query1=("SELECT Quantity FROM ITEMS Where Item = ?")
            cursor.execute(query1,MyTuple)
            result = cursor.fetchone()
            if result is not None:
                
                current_qty = float(result[0])
                new_qty=current_qty-float(IQuantity)
                if(new_qty<0):
                    msg1="Unable to Delete. Items to be deleted exceeds total Item Limit."
                    messagebox.showinfo("Item Not Deleted", msg1)
                    return

                query2=("UPDATE ITEMS SET Quantity = ? WHERE Item = ?")
                cursor.execute(query2,(new_qty,IName))
                
                
                #query = ("DELETE FROM ITEMS WHERE Item = ?")
                query = ("")
                #cursor.execute(query,MyTuple)
                conn.commit()
                msg1=IName+" deleted from database."
                messagebox.showinfo("Item Deleted", msg1)
                initialize_item()
            #self.showall()
            self.entry3f.delete(0, tk.END)
            self.entry3h.delete(0, tk.END)
            self.entry3j.delete(0, tk.END)
            # Close the cursor
            cursor.close()
            initialize_item()
            self.view_item()
        except sqlite3.Error as error:
            print('Error occurred - ',error)

    def update(self):
        try:    
            #conn = sqlite3.connect('firsttable.db')
            conn = sqlite3.connect('secondtable.db')
            cursor = conn.cursor()

            IName = self.entry4a.get()
            IPrice = self.entry4c.get()
            IPrice = float(IPrice)
            mylist=[]
            mylist.append(IPrice)
            mylist.append(IName)
            MyTuple=tuple(mylist)
            mylist.clear()

            query = "UPDATE ITEMS SET Price = ? WHERE Item = ?"
            
            cursor.execute(query,MyTuple)
            conn.commit()

            IPrice =str(IPrice)
            msg1=IName+" price updated to "+IPrice
            messagebox.showinfo("Item Deleted", msg1)

            self.entry4a.delete(0, tk.END)
            self.entry4b.delete(0, tk.END)
            self.entry4c.delete(0, tk.END)
            
            self.showall()
            # Close the cursor
            cursor.close()
            self.view_item()
        except sqlite3.Error as error:
            print('Error occurred - ',error)

    def set_price(self,event):
        item = self.desc_entry.get()
        #conn = sqlite3.connect('firsttable.db')
        conn = sqlite3.connect('secondtable.db')
        cursor = conn.cursor()
        query = ("SELECT Price FROM ITEMS WHERE Item = ?")
        cursor.execute(query, (item,))
        result = cursor.fetchone()
        if result:
            self.price_spinbox.delete(0, tk.END)
            self.price_spinbox.insert(0, str(result[0]))
        cursor.close()

    def set_price1(self,event):
        item = self.entry4a.get()
        #conn = sqlite3.connect('firsttable.db')
        conn = sqlite3.connect('secondtable.db')
        cursor = conn.cursor()
        query = ("SELECT Price FROM ITEMS WHERE Item = ?")
        cursor.execute(query, (item,))
        result = cursor.fetchone()
        if result:
            self.entry4b.delete(0, tk.END)
            self.entry4b.insert(0, str(result[0]))
        cursor.close()

    def set_price2(self,event):
        item = self.entry3f.get()
        #conn = sqlite3.connect('firsttable.db')
        conn = sqlite3.connect('secondtable.db')
        cursor = conn.cursor()
        query = ("SELECT Price,Quantity FROM ITEMS WHERE Item = ?")
        cursor.execute(query, (item,))
        result = cursor.fetchone()
        if result:
            self.entry3h.delete(0, tk.END)
            self.entry3h.insert(0, str(result[0]))
            self.entry3j.insert(0, str(result[1]))
        cursor.close()
        
    def remove_item(self):
        selected_items = self.tree.selection()
        for item_id in selected_items:
            item_values = self.tree.item(item_id, "values")
            qty = item_values[0]
            desc = item_values[1]
            items_to_remove = []
            for item in invoice_list:
                if str(item[0]) == str(qty) and str(item[1]) == str(desc):
                    invoice_list.remove(item)
                    break

            self.tree.delete(item_id)  # Delete item from the Treeview
        self.calculate_total()
        #for it in invoice_list:
        self.calculate_total()

    def update_date(self):
        current_date_time = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
        self.date_entry.delete(0, tk.END)  # Clear the current value
        self.date_entry.insert(0, current_date_time)

    def schedule_date_update(self):
        # Schedule the update_date function to be called every 30 seconds (30000 milliseconds)
        self.after(30000, self.schedule_date_update)
        self.update_date()
        

    # Load the sequence number from the database
    def load_sequence_number(self):
        conn = sqlite3.connect('secondtable.db')
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM SequenceNumber')
        row = cursor.fetchone()
        conn.close()

        if row:
            return {'last_date': row[1], 'number': row[2]}
        else:
            return {'last_date': None, 'number': 1}

    # Save the sequence number to the database
    def save_sequence_number(self,seq_number):
        conn = sqlite3.connect('secondtable.db')
        cursor = conn.cursor()
        cursor.execute('UPDATE SequenceNumber SET last_date=?, number=?', (seq_number['last_date'], seq_number['number']))
        conn.commit()
        conn.close()

    def get_invoice_number(self):
        global invoicenumber
        prefix='INV-'
        # Get current date
        current_date = datetime.date.today()
        
        # Load sequence number from the database
        self.connect()
        seq_number = self.load_sequence_number()
        current_timestamp = datetime.datetime.now().strftime('%y%m%d')
        seq_number_str = str(seq_number['number']).zfill(4)

        # Concatenate prefix, timestamp, and sequential number
        invoice_number = f'{prefix}{current_timestamp}-{seq_number_str}'
        invoicenumber = invoicenumber

        return invoice_number
        
        
    def generate_invoice_number(self):
        global invoicenumber
        prefix='INV-'
        # Get current date
        current_date = datetime.date.today()
        
        # Load sequence number from the database
        seq_number = self.load_sequence_number()
        last_date = datetime.datetime.strptime(seq_number['last_date'], '%Y-%m-%d').date()

        # If current date is different from the last saved date, reset sequence number
        
        if current_date != last_date:
            seq_number['last_date'] = current_date
            seq_number['number'] = 1
        else:
            seq_number['number'] += 1
    
        # Save updated sequence number to the database
        
        self.save_sequence_number(seq_number)

        # Get current timestamp
        current_timestamp = datetime.datetime.now().strftime('%y%m%d')

        # Format sequential number with leading zeros
        seq_number_str = str(seq_number['number']).zfill(4)

        # Concatenate prefix, timestamp, and sequential number
        invoice_number = f'{prefix}{current_timestamp}-{seq_number_str}'
        invoicenumber = invoicenumber

        return invoice_number

if __name__ == "__main__":
    app = MultiFrameApp()
    app.mainloop()
