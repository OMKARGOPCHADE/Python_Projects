from tkinter import *
import tkinter as tk
from tkinter import ttk
from database import *
from tkinter import messagebox as ms
import win32api
import win32print
from tkinter import filedialog
from reportlab.pdfgen import canvas
import xlsxwriter
from database import *
from generate_qr import *
class Student:
    def __init__(self,root):
        self.root = root
        self.root.title("Student Management System")
        self.root.geometry("1350x700+0+0")
        title = Label(self.root,text="Student Management System",bd=8,relief=GROOVE,font=("times new roman",40,"bold"),bg="yellow",fg="red")
        title.pack(side=TOP,fill=X)

        #=== for taking user data variables. ===#
        self.Roll_No = StringVar()
        self.Name = StringVar()
        self.Email = StringVar()
        self.Age = StringVar()
        self.Gender = StringVar()
        self.Contact = StringVar()
        self.Search = StringVar()
        self.txt_search = StringVar()
#### Manage Frame #####
        Manage_Frame = Frame(self.root,bd=4,relief=RIDGE,bg="crimson")
        Manage_Frame.place(x=20,y=100,width=450,height=580)

        m_title = Label(Manage_Frame,text="Manage Students",bg="crimson",fg="white",font=("times new roman",30,"bold"))
        m_title.grid(row=0,columnspan=2,pady=20)
        
        lbl_roll = Label(Manage_Frame,text="Roll No: ",bg="crimson",fg="white",font=("times new roman",20,"bold"))
        lbl_roll.grid(row=1,column=0,pady=10,padx=20,sticky="w")
        
        text_roll = Entry(Manage_Frame,textvariable=self.Roll_No,font=("times new roman",15,"bold"),bd=5,relief=GROOVE)
        text_roll.grid(row=1,column=1,pady=10,padx=20,sticky="w")

        lbl_name = Label(Manage_Frame,text="Name: ",bg="crimson",fg="white",font=("times new roman",20,"bold"))
        lbl_name.grid(row=2,column=0,pady=10,padx=20,sticky="w")
        
        text_name = Entry(Manage_Frame,textvariable=self.Name,font=("times new roman",15,"bold"),bd=5,relief=GROOVE)
        text_name.grid(row=2,column=1,pady=10,padx=20,sticky="w")

        lbl_email = Label(Manage_Frame,text="Email: ",bg="crimson",fg="white",font=("times new roman",20,"bold"))
        lbl_email.grid(row=3,column=0,pady=10,padx=20,sticky="w")
        
        text_email = Entry(Manage_Frame,textvariable=self.Email,font=("times new roman",15,"bold"),bd=5,relief=GROOVE)
        text_email.grid(row=3,column=1,pady=10,padx=20,sticky="w")

        lbl_age = Label(Manage_Frame,text="Age: ",bg="crimson",fg="white",font=("times new roman",20,"bold"))
        lbl_age.grid(row=4,column=0,pady=10,padx=20,sticky="w")
        
        text_age = Entry(Manage_Frame,textvariable=self.Age,font=("times new roman",15,"bold"),bd=5,relief=GROOVE)
        text_age.grid(row=4,column=1,pady=10,padx=20,sticky="w")

        lbl_gender = Label(Manage_Frame,text="Gender: ",bg="crimson",fg="white",font=("times new roman",20,"bold"))
        lbl_gender.grid(row=5,column=0,pady=10,padx=20,sticky="w")

        combo_gender = ttk.Combobox(Manage_Frame,textvariable=self.Gender,font=("times new roman",13,"bold"),state='readonly')
        combo_gender['values']=("Male","Female","Other")
        combo_gender.grid(row=5,column=1,pady=10,padx=20,sticky="w")

        lbl_contact = Label(Manage_Frame,text="Contact No: ",bg="crimson",fg="white",font=("times new roman",20,"bold"))
        lbl_contact.grid(row=6,column=0,pady=10,padx=20,sticky="w")
        
        text_contact = Entry(Manage_Frame,textvariable=self.Contact,font=("times new roman",15,"bold"),bd=5,relief=GROOVE)
        text_contact.grid(row=6,column=1,pady=10,padx=20,sticky="w")

        lbl_address = Label(Manage_Frame,text="Address: ",bg="crimson",fg="white",font=("times new roman",20,"bold"))
        lbl_address.grid(row=7,column=0,pady=10,padx=20,sticky="w")

        self.text_address = Text(Manage_Frame,width=25,height=4)
        self.text_address.grid(row=7,column=1,pady=10,padx=20,sticky="w")  
        #self.text_address.get('1.0',END)
        # ==== Button Frame ====
        btn_Frame = Frame(Manage_Frame,bd=4,relief=RIDGE,bg="crimson")
        btn_Frame.place(x=15,y=517,width=420)

        add_btn = Button(btn_Frame,command=self.add,text="Add",width=10).grid(row=0,column=0,padx=10,pady=10)
        upd_btn = Button(btn_Frame,command=self.update,text="Update",width=10).grid(row=0,column=1,padx=10,pady=10)
        del_btn = Button(btn_Frame,command=self.delete,text="Delete",width=10).grid(row=0,column=2,padx=10,pady=10)
        clr_btn = Button(btn_Frame,command=self.clear,text="Clear",width=10).grid(row=0,column=3,padx=10,pady=10)

#### Details Frame #####
        Details_Frame = Frame(self.root,bd=4,relief=RIDGE,bg="crimson")
        Details_Frame.place(x=500,y=100,width=950,height=580)

        lbl_search = Label(Details_Frame,text="Search by",bg="crimson",fg="white",font=("times new roman",20,"bold"))
        lbl_search.grid(row=0,column=0,pady=10,padx=20,sticky="w")

        combo_search = ttk.Combobox(Details_Frame,textvariable=self.Search,width=10,font=("times new roman",13,"bold"),state='readonly')
        combo_search['values']=("Roll No","Name","Contact No","Gender")
        combo_search.grid(row=0,column=1,pady=10,padx=20,sticky="w")

        text_search = Entry(Details_Frame,textvariable=self.txt_search,width=15,font=("times new roman",15,"bold"),bd=5,relief=GROOVE)
        text_search.grid(row=0,column=2,pady=10,padx=20,sticky="w")

        search_btn = Button(Details_Frame,command=self.search,text="Search",width=10,pady=5).grid(row=0,column=3,padx=10,pady=10)
        showall_btn = Button(Details_Frame,command=self.fetch,text="Show All",width=10,pady=5).grid(row=0,column=4,padx=10,pady=10)
        print_pdf = Button(Details_Frame,command=self.save_as_pdf,text="Print",width=6,pady=4).grid(row=0,column=5,padx=10,pady=10)
        print_xl = Button(Details_Frame,command=self.print_xl,text="Print_xl",width=6,pady=4).grid(row=0,column=6,padx=10,pady=10)
        gen_qr = Button(Details_Frame,command=self.gen_qr,text="Gen_QR",width=6,pady=4).grid(row=0,column=7,padx=10,pady=10)
#=== Table Frame ===#
        Record_Frame = Frame(Details_Frame,bd=4,relief=RIDGE,bg="crimson")
        Record_Frame.place(x=10,y=70,width=760,height=500)    

        x_scroll = Scrollbar(Record_Frame,orient=HORIZONTAL)
        y_scroll = Scrollbar(Record_Frame,orient=VERTICAL)
        self.S_table = ttk.Treeview(Record_Frame,columns=("Roll","Name","Email","Age","Gender","Contact","Address"),xscrollcommand=x_scroll.set,yscrollcommand=y_scroll.set)
        x_scroll.pack(side=BOTTOM,fill=X)
        y_scroll.pack(side=RIGHT,fill=Y)
        x_scroll.config(command=self.S_table.xview)
        y_scroll.config(command=self.S_table.yview)
        self.S_table.heading("Roll",text="Roll No")
        self.S_table.heading("Name",text="Name")
        self.S_table.heading("Email",text="Email")
        self.S_table.heading("Age",text="Age")
        self.S_table.heading("Gender",text="Gender")
        self.S_table.heading("Contact",text="Contact")
        self.S_table.heading("Address",text="Address")
        self.S_table['show']='headings'
        self.S_table.column("Roll",width=80)
        self.S_table.column("Name",width=100)
        self.S_table.column("Email",width=120)
        self.S_table.column("Age",width=80)
        self.S_table.column("Gender",width=100)
        self.S_table.column("Contact",width=100)
        self.S_table.column("Address",width=170)
        self.S_table.pack(fill=BOTH,expand=1)
        self.S_table.bind("<ButtonRelease-1>",self.get_cursor)
        self.fetch()
        ### Functions ###
        ### add function ###
    def add(self):
         if self.Roll_No.get()=="" or self.Name.get()=="" or self.Age.get()=="" or self.Email.get()=="" or self.Contact.get()=="" or self.text_address.get('1.0',END)=="":
              ms.showerror("Error","All fields are required!!!")
         else :   
               if self.validate_roll() and self.validate_contact():         
                    obj = dataset();
                    obj.add_data(int(self.Roll_No.get()),self.Name.get(),self.Email.get(),
                                 int(self.Age.get()),self.Gender.get(),int(self.Contact.get()),
                                 self.text_address.get('1.0',END))
                    self.fetch()
                    self.clear()
                    obj.close_connection()
                    ms.showinfo("Success","Recored inserted successfully!")
        ### retrive data function ###
    def fetch(self):
         obj = dataset()
         rows = obj.retrive_data()
         if len(rows)!=0:
              self.S_table.delete(*self.S_table.get_children())
              for row in rows:
                   self.S_table.insert('',END,values=row)
              self.Search.set("")
              self.txt_search.set("")          
         obj.close_connection()
        ### clear function ###
    def clear(self):
         self.Roll_No.set("")
         self.Name.set("")
         self.Email.set("")
         self.Age.set("")
         self.Gender.set("")
         self.Contact.set("")
         self.text_address.delete('1.0',END)
        ### Update Function ###
    def update(self):
         obj = dataset()
         search = int(self.Roll_No.get())
         rows = obj.search_roll_data(search)
         if len(rows) != 0:
               obj.update_data(int(self.Roll_No.get()),self.Name.get(),self.Email.get(),
                               int(self.Age.get()),self.Gender.get(),int(self.Contact.get()),
                               self.text_address.get('1.0',END))
               self.fetch()
               self.clear()
         else :
             ms.showerror("Error","Data no not present u can'not update absent data!")
         obj.close_connection()
        ### Delete Function ###
    def delete(self):
         obj = dataset()
         search = int(self.Roll_No.get())
         rows = obj.search_roll_data(search)
         if len(rows) != 0:
               obj.delete_data(int(self.Roll_No.get()))
               self.fetch()
               self.clear()
         else :
             ms.showerror("Error","Data no not present u can'not delete absent data!")
         obj.close_connection()
    def search(self):
         obj = dataset()
         ### search with roll no ###
         if(self.Search.get() == "Roll No"):
              search= (int(self.txt_search.get()) )
              rows = obj.search_roll_data(search)
              if len(rows)!=0:
                   self.S_table.delete(*self.S_table.get_children())
                   for row in rows:
                        self.S_table.insert('',END,values=row)
              else :
                   ms.showerror("Error","Roll no not present!")
              self.Search.set("")
              self.txt_search.set("")
              obj.close_connection()
         ### search with name ###
         if(self.Search.get() == "Name"):
              search= (self.txt_search.get())
              rows = obj.search_name_data(search)
              if len(rows)!=0:
                   self.S_table.delete(*self.S_table.get_children())
                   for row in rows:
                        self.S_table.insert('',END,values=row)
              else :
                   ms.showerror("Error","Name not present!")
              self.Search.set("")
              self.txt_search.set("")
              obj.close_connection()
         ### search with contact ###
         if(self.Search.get() == "Contact No"):
              search= (int(self.txt_search.get()))
              rows = obj.search_contact_data(search)
              if len(rows)!=0:
                   self.S_table.delete(*self.S_table.get_children())
                   for row in rows:
                        self.S_table.insert('',END,values=row)
              else :
                   ms.showerror("Error","Contact not present!")
              self.Search.set("")
              self.txt_search.set("")                   
              obj.close_connection()
         ### search with Gender ###
         if(self.Search.get() == "Gender"):
              search= (self.txt_search.get())
              rows = obj.search_gender_data(search)
              if len(rows)!=0:
                   self.S_table.delete(*self.S_table.get_children())
                   for row in rows:
                        self.S_table.insert('',END,values=row)
              else :
                   ms.showerror("Error","Gender not present!")
              self.Search.set("")
              self.txt_search.set("")                   
              obj.close_connection()
          ### function for cursor ###
    def get_cursor(self,ev):
         cursor_row = self.S_table.focus()
         contents = self.S_table.item(cursor_row)
         row = contents['values']
         self.Roll_No.set(row[0])
         self.Name.set(row[1])
         self.Email.set(row[2])
         self.Age.set(row[3])
         self.Gender.set(row[4])
         self.Contact.set(row[5])
         self.text_address.delete('1.0',END)
         self.text_address.insert(END,row[6])
     ### Validation for mobile number Function ###
    def validate_roll(self):
         obj = dataset()
         search = int(self.Roll_No.get())
         rows = obj.search_roll_data(search)
         if len(rows) == 0:
              return True
         else :
               ms.showerror("Error","Roll no must be different!!")
         obj.close_connection()
     ### Validation for contact no Function ###
    def validate_contact(self):
         obj = dataset()
         search= (self.Contact.get())
         if len(search) != 10:
              ms.showerror("Error","Enter Valid Contact Number!!")
              return False
         rows = obj.search_contact_data(int(search))
         if len(rows) != 0:
               ms.showerror("Error","Contact no must be different!!")
         else :
              return True
         obj.close_connection()
     ### Printing Data Function ###
    def save_as_pdf(self):
      # Get data from Treeview
      data = []
      columns = ["Roll","Name","Email","Age","Gender","Contact","Address"]
     #  data.append(columns)  # Add header row
      for item in self.S_table.get_children():
        data.append(self.S_table.item(item, "values"))
    
      # Create PDF with reportlab
      filename = "treeview_data.pdf"
      pdf = canvas.Canvas(filename, pagesize=(595.28, 841.89))
      pdf.setFont("Helvetica", 8)
      pdf.translate(10,800)  
      column_width = 80  
      x, y = 0, 0
      for col_name in columns:
        pdf.drawString(x, y, col_name)
        if column_width:  
          x += column_width
        else:
          x += len(col_name) * 7  
      column_width = 80
      y -= 20  
      for row in data:
        x = 0
        for value in row:
          pdf.drawString(x, y, str(value))
          if column_width:  
            x += column_width
          else:
            x += len(str(value)) * 7  
        y -= 15  
    
      pdf.save()
      
      tk.messagebox.showinfo("Success", f"Data saved to '{filename}'")
     #creating xl sheet
    def print_xl(self):
          filename = "Student_SE_IT.xlsx"
          workbook = xlsxwriter.Workbook(filename)
          worksheet = workbook.add_worksheet("firstsheet")

          worksheet.write(0, 0, "#")
          worksheet.write(0, 1, "Roll No")
          worksheet.write(0, 2, "Name")
          worksheet.write(0, 3, "Email")
          worksheet.write(0, 4, "Age")
          worksheet.write(0, 5, "Gender")
          worksheet.write(0, 6, "Mobile No")
          worksheet.write(0, 7, "Address")

          obj = dataset()
          rows = obj.retrive_data()

          if len(rows) != 0:
              for index, row in enumerate(rows, start=1):
                  worksheet.write(index, 0, str(index - 1))  # Index starts from 1 in enumerate
                  worksheet.write(index, 1, row[0])
                  worksheet.write(index, 2, row[1])
                  worksheet.write(index, 3, row[2])
                  worksheet.write(index, 4, row[3])
                  worksheet.write(index, 5, row[4])
                  worksheet.write(index, 6, row[5])
                  worksheet.write(index, 7, row[6])

          workbook.close()
          tk.messagebox.showinfo("Success", f"Data saved to '{filename}'")
          obj.close_connection() 
    def gen_qr(self):
               filename = "Student_IT.png"
               qrCode(filename)  
               tk.messagebox.showinfo("Success", f"QR Code Genrated to '{filename}'")  
root = Tk()
ob = Student(root)
root.mainloop()