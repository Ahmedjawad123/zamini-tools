from tkinter import *
from tkinter import messagebox
import os
import app_info  # Make sure app_info.py is in the same folder

if not os.path.exists('file'):
    os.makedirs('file')

class Login:
    def __init__(self, root):
        self.root = root
        self.root.title(f"Zamini File Manager v{app_info.VERSION}")
        self.root.geometry('1350x700+0+0')

        # Set window icon (use your .ico file, convert PNG to ICO if needed)
        try:
            self.root.iconbitmap('Zamini_Musafir_logo.ico')
        except Exception as e:
            print("Icon file not found or invalid:", e)

        # Create menu bar with Help menu
        menu_bar = Menu(self.root)
        help_menu = Menu(menu_bar, tearoff=0)
        help_menu.add_command(label="How to Use", command=lambda: messagebox.showinfo("Help", app_info.HELP_TEXT))
        help_menu.add_command(label="About Us", command=lambda: messagebox.showinfo("About", app_info.ABOUT_TEXT))
        menu_bar.add_cascade(label="Help", menu=help_menu)
        self.root.config(menu=menu_bar)

        loginframe = Frame(self.root, bd=10, relief=GROOVE)
        loginframe.place(x=450, y=150, height=300)

        self.var_user = StringVar()
        self.var_password = StringVar()

        lbl_title = Label(loginframe, text='Login System', font=('times new roman', 30, 'bold'))
        lbl_title.grid(row=0, column=0, columnspan=2)

        lbl_username = Label(loginframe, text='User Name', font=('times new roman', 15, 'bold'))
        lbl_username.grid(row=1, column=0, padx=10, pady=10, sticky='w')
        txt_username = Entry(loginframe, textvariable=self.var_user, bd=4, font=('times new roman', 15, 'bold'), width=20)
        txt_username.grid(row=1, column=1, padx=10, pady=10)

        lbl_password = Label(loginframe, text='Password', font=('times new roman', 15, 'bold'))
        lbl_password.grid(row=2, column=0, padx=10, pady=10, sticky='w')
        txt_password = Entry(loginframe, show='*', textvariable=self.var_password, bd=4, font=('times new roman', 15, 'bold'), width=20)
        txt_password.grid(row=2, column=1, padx=10, pady=10)

        btn_login = Button(loginframe, bd=4, text='Login', font=('times new roman', 15, 'bold'), width=7,
                           command=self.login)
        btn_login.place(x=0, y=180)
        btn_clear = Button(loginframe, bd=4, text='Clear', font=('times new roman', 15, 'bold'), width=7,
                           command=self.clear)
        btn_clear.place(x=125, y=180)
        btn_exit = Button(loginframe, bd=4, text='Exit', font=('times new roman', 15, 'bold'), width=7,
                          command=self.exit)
        btn_exit.place(x=250, y=180)

        # Info label with username and password hint
        info_label = Label(loginframe, 
                           text="(Hint: Username = Ahmed, Password = 1234)", 
                           font=('times new roman', 12, 'italic'), 
                           fg='gray')
        info_label.place(x=0, y=230)

    def login(self):
        if self.var_user.get() == '' or self.var_password.get() == '':
            messagebox.showerror('Error', 'All Fields are required')
        elif self.var_user.get() == 'Ahmed' and self.var_password.get() == '1234':
            messagebox.showinfo('Success', f'Welcome {self.var_user.get()}\nPassword {self.var_password.get()}')
            self.root.destroy()
            import File_Based_main
            File_Based_main.file_app()
        else:
            messagebox.showerror('Error', 'Username or Password is invalid')

    def clear(self):
        self.var_user.set('')
        self.var_password.set('')

    def exit(self):
        option = messagebox.askyesno('Exit', 'Do you really want to exit?')
        if option:
            self.root.destroy()

if __name__ == "__main__":
    root = Tk()
    obj = Login(root)
    root.mainloop()
