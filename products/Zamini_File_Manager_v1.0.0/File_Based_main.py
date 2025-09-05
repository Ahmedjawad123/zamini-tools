from tkinter import *
from tkinter import ttk, messagebox
import time
import os
import app_info  # Make sure app_info.py is in the same folder

if not os.path.exists('file'):
    os.makedirs('file')

class file_app:
    def __init__(self):
        self.root = Tk()
        self.root.title('File Based Record System')
        self.root.geometry('1350x700+0+0')

        # --- Add Menu Bar with Help Menu ---
        menu_bar = Menu(self.root)
        help_menu = Menu(menu_bar, tearoff=0)
        help_menu.add_command(label="How to Use", command=lambda: messagebox.showinfo("Help", app_info.HELP_TEXT))
        help_menu.add_command(label="About Us", command=lambda: messagebox.showinfo("About", app_info.ABOUT_TEXT))
        menu_bar.add_cascade(label="Help", menu=help_menu)
        self.root.config(menu=menu_bar)
        # ------------------------------------

        lbl_title = Label(self.root, text='File Based Record System', font=('times new roman', 28, 'bold'), bd=10, relief=GROOVE, pady=10).pack(fill=X)
        frame_student = Frame(self.root, bd=10, relief=GROOVE, pady=10)
        frame_student.place(x=10, y=80, width=900, height=400)
        lbl_title_student = Label(frame_student, text='Student ID', font=('times new roman', 24, 'bold')).grid(row=0, columnspan=4, pady=(0, 20))

        self.var_s_id = StringVar()
        self.var_name = StringVar()
        self.var_course = StringVar()
        self.var_address = StringVar()
        self.var_city = StringVar()
        self.var_contact = StringVar()
        self.var_date = StringVar()
        self.var_degree = StringVar()
        self.var_id_prof = StringVar()
        self.var_payment = StringVar()

        lbl_id = Label(frame_student, text='Student ID', font=('times new roman', 18, 'bold')).grid(row=1, column=0, padx=10, pady=10, sticky='w')
        txt_id = Entry(frame_student, textvariable=self.var_s_id, font=('times new roman', 18, 'bold'), width=17, bd=4).grid(row=1, column=1, padx=10, pady=10)
        lbl_name = Label(frame_student, text='Name', font=('times new roman', 18, 'bold')).grid(row=2, column=0, padx=10, pady=10, sticky='w')
        txt_name = Entry(frame_student, textvariable=self.var_name, font=('times new roman', 18, 'bold'), width=17, bd=4).grid(row=2, column=1, padx=10, pady=10)
        lbl_course = Label(frame_student, text='Course', font=('times new roman', 18, 'bold')).grid(row=3, column=0, padx=10, pady=10, sticky='w')
        txt_course = Entry(frame_student, textvariable=self.var_course, font=('times new roman', 18, 'bold'), width=17, bd=4).grid(row=3, column=1, padx=10, pady=10)
        lbl_address = Label(frame_student, text='Address', font=('times new roman', 18, 'bold')).grid(row=4, column=0, padx=10, pady=10, sticky='w')
        txt_address = Entry(frame_student, textvariable=self.var_address, font=('times new roman', 18, 'bold'), width=17, bd=4).grid(row=4, column=1, padx=10, pady=10)
        lbl_city = Label(frame_student, text='City', font=('times new roman', 18, 'bold')).grid(row=5, column=0, padx=10, pady=10, sticky='w')
        txt_city = Entry(frame_student, textvariable=self.var_city, font=('times new roman', 18, 'bold'), width=17, bd=4).grid(row=5, column=1, padx=10, pady=10)

        lbl_contact = Label(frame_student, text='Contact No:', font=('times new roman', 18, 'bold')).grid(row=1, column=2, padx=10, pady=10, sticky='w')
        txt_contact = Entry(frame_student, textvariable=self.var_contact, font=('times new roman', 18, 'bold'), width=17, bd=4).grid(row=1, column=3, padx=10, pady=10)
        lbl_date = Label(frame_student, text='Date', font=('times new roman', 18, 'bold')).grid(row=2, column=2, padx=10, pady=10, sticky='w')
        txt_date = Entry(frame_student, textvariable=self.var_date, font=('times new roman', 18, 'bold'), width=17, bd=4).grid(row=2, column=3, padx=10, pady=10)
        lbl_degree = Label(frame_student, text='Select Degree', font=('times new roman', 18, 'bold')).grid(row=3, column=2, padx=10, pady=10, sticky='w')
        combotxt_degree = ttk.Combobox(frame_student, textvariable=self.var_degree, font=('times new roman', 16, 'bold'), width=13, state='readonly')
        combotxt_degree['values'] = ('BCA', 'BBA', 'FA', 'I.Com', 'Metrc',)
        combotxt_degree.grid(row=3, column=3, padx=10, pady=10, sticky='w')
        lbl_id_prof = Label(frame_student, text='ID Proof', font=('times new roman', 18, 'bold')).grid(row=4, column=2, padx=10, pady=10, sticky='w')
        combotxt_id_prof = ttk.Combobox(frame_student, textvariable=self.var_id_prof, font=('times new roman', 16, 'bold'), width=13, state='readonly')
        combotxt_id_prof['values'] = ('School Card', 'Bank Card', 'Passport', 'National Card', 'Driving Card')
        combotxt_id_prof.grid(row=4, column=3, padx=10, pady=10, sticky='w')
        lbl_payment = Label(frame_student, text='Payment Method', font=('times new roman', 18, 'bold')).grid(row=5, column=2, padx=10, pady=10, sticky='w')
        combotxt_payment = ttk.Combobox(frame_student, textvariable=self.var_payment, font=('times new roman', 16, 'bold'), width=13, state='readonly')
        combotxt_payment['values'] = ('Bank Transfer', 'Cash', 'Card', 'Internet Banking')
        combotxt_payment.grid(row=5, column=3, padx=10, pady=10, sticky='w')




        frame_btn=Frame(self.root,bd=10,relief=GROOVE,pady=10)
        frame_btn.place(x=10,y=480)
        
        btn_save=Button(frame_btn,text='Save',font=('times new roman',18,'bold'),width=15,bd=3,
                        command=self.save_file).grid(row=0,column=0,padx=10,pady=10)
        btn_delete=Button(frame_btn,text='Delete',font=('times new roman',18,'bold'),width=15,bd=3,
                          command=self.delete).grid(row=0,column=1,padx=10,pady=10)
        btn_new=Button(frame_btn,text='New',font=('times new roman',18,'bold'),width=15,bd=3,
                       command=self.new).grid(row=0,column=2,padx=10,pady=10)
        btn_logout=Button(frame_btn,text='Log Out',font=('times new roman',18,'bold'),width=15,bd=3,
                          command=self.logout).grid(row=0,column=3,padx=10,pady=10)
        btn_exit=Button(frame_btn,text='Exit',font=('times new roman',18,'bold'),width=15,bd=3,
                        command=self.exit).grid(row=0,column=4,padx=10,pady=10)
        
        frame_file=Frame(self.root,bd=10,relief=GROOVE,pady=10)
        frame_file.place(x=910,y=80,width=340,height=400)
        lbl_title_student=Label(frame_file,text='All Files',font=('times new roman',22,'bold'),bd=5,relief=GROOVE,pady=10).pack(side=TOP,fill=X)
        scroll_y=Scrollbar(frame_file,orient=VERTICAL)
        self.file_list=Listbox(frame_file,font=('times new roman',12,'bold'),yscrollcommand=scroll_y.set)
        scroll_y.pack(fill=Y,side=RIGHT)
        scroll_y.config(command=self.file_list.yview)
        self.file_list.pack(fill=BOTH,expand=1)
        self.file_list.bind('<ButtonRelease-1>',self.get_data)
        self.show_data()
        self.root.mainloop()

        
        
    def save_file(self):
        present='no'
        if self.var_s_id.get()=='':
            messagebox.showerror('Error','Student Id must be required')
        else:
            f=os.listdir('file/')
            for i in f:
                if i.split('.')[0]==self.var_s_id.get():
                    present='yes'
            if present=='yes':
                ask=messagebox.askyesno('Update','File already present.\nDo you want to update update?')
                if ask>0:
                    self.save_data()
                    messagebox.showinfo('Success',f'Your file number: {self.var_s_id.get()}\nSuccessfully saved')
                    self.show_data()
            else:
                self.save_data()
                messagebox.showinfo('Success',f'Your file number: {self.var_s_id.get()}\nSuccessfully saved')
                self.show_data()

        if  self.var_s_id.get()=='' or self.var_name.get()=='' or self.var_course.get()=='' or self.var_address.get()=='' or self.var_city.get()=='' or self.var_contact.get()=='' or self.var_date.get()=='' or self.var_degree.get()=='' or self.var_id_prof.get()=='' or self.var_payment.get()=='':
            messagebox.showerror('Error','All Fields are required')
        else:
            self.save_data()
    def save_data(self):
        f=open('file/'+str(self.var_s_id.get())+'.txt','w')
        f.write(
                str(self.var_s_id.get())+','+
                str(self.var_name.get())+','+
                str(self.var_course.get())+','+
                str(self.var_address.get())+','+
                str(self.var_city.get())+','+
                str(self.var_contact.get())+','+
                str(self.var_date.get())+','+
                str(self.var_degree.get())+','+
                str(self.var_id_prof.get())+','+
                str(self.var_payment.get())
                )
        f.close()
        

    def show_data(self):
        files=os.listdir('file/')
        self.file_list.delete(0,END)
        for i in files:
            self.file_list.insert(END,i)
     
    def get_data(self,ev):
        data_cursor=self.file_list.curselection()
        f1=open('file/'+str(self.file_list.get(data_cursor)))
        value=[]
        for f in f1:
            value=f.split(',')
        self.var_s_id.set(value[0])
        self.var_name.set(value[1])
        self.var_course.set(value[2])
        self.var_address.set(value[3])
        self.var_city.set(value[4])
        self.var_contact.set(value[5])
        self.var_date.set(value[6])
        self.var_degree.set(value[7])
        self.var_id_prof.set(value[8])
        self.var_payment.set(value[9])
    def new(self):
        self.var_s_id.set('')
        self.var_name.set('')
        self.var_course.set('')
        self.var_address.set('')
        self.var_city.set('')
        self.var_contact.set('')
        self.var_date.set('')
        self.var_degree.set('')
        self.var_id_prof.set('')
        self.var_payment.set('')
    
    def delete(self):
        present='no'
        if self.var_s_id.get()=='':
            messagebox.showerror('Error','Student Id must be required')
        else:
            f=os.listdir('file/')
            for i in f:
                if i.split('.')[0]==self.var_s_id.get():
                    present='yes'
            if present=='yes':
                ask=messagebox.askyesno('Delete','Are you sure you want to Delete?')
                if ask>0:
                    os.remove('file/'+str(self.var_s_id.get())+'.txt')
                    messagebox.showinfo('Deleted',f'Your file number: {self.var_s_id.get()}\nSuccessfully Deleted')
                    self.show_data()

                else:
                    messagebox.showerror('Error','File Not Found')
    def exit(self):
        ask=messagebox.askyesno('Exit','Do you really want to Exit?')
        if ask>0:
            self.root.destroy()    

    def logout(self):
        ask = messagebox.askyesno('Logout', 'Do you want to logout?')
        if ask:
            self.root.destroy()  # close current window

            import tkinter as tk
            import File_Based_Login

            new_root = tk.Tk()
            File_Based_Login.Login(new_root)
            new_root.mainloop()


        
    
    
    
     
        
