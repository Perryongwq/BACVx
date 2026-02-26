from tkinter import *

from Utils.readSettings import readSettings
from Pages.Settings import Settings

class Login(readSettings):
        
    """
    Tkinter Component with Simple Login Page
    Parameters
    ----------
    root : root
        The base root.
    cap : VideoCapture
        Returns a tuple (success, frame)
    Wscreen : str
        Providing Screen Size (Width) to Class.
    Hscreen : str
        Providing Screen Size (Height) to Class.
    light : Lighting Component
        Instantiate Lighting Component to pass to next Class
    """
        
    def __init__(self,root,cap,Wscreen,Hscreen,light,mat):
        super().__init__()
        self.root = Toplevel(root)
        self.win_config(Wscreen,Hscreen)
        self.widgets(root,cap,Wscreen,light,mat)
        self.root.grab_set()

    def win_config(self,Wscreen,Hscreen):
        self.root.title("ME Login")
        self.root.config(bg='lightgrey',bd=50)
        self.root.geometry(f"{int(Wscreen/4)}x{int(Hscreen/4)}+{int(Wscreen*3/8)}+{int(Hscreen*3/8)}")
        self.root.rowconfigure(index=5,weight=1)
        self.root.columnconfigure(index=1, weight=1)
        self.root.columnconfigure(index=4, weight=1)

    def widgets(self,root,cap,Wscreen,light,mat):

        # StringVars (used to retrieve input from entry)
        self.the_user = StringVar()
        self.the_pass = StringVar()

        self.LogEntry = {}
        self.LogEntry['Username'] = Entry(self.root,textvariable=self.the_user)
        self.LogEntry['Password'] = Entry(self.root,textvariable=self.the_pass,show='*')
        self.bad_pass = Label(self.root,bg='red')

        for i,cred in enumerate(self.credentials):
            Label(self.root,text=cred+' :',background='lightgrey').grid(row=i,column=1)
            self.LogEntry[cred].grid(row=i,column=2,columnspan=2)
        
        Button(self.root,text="Login",command=lambda:self.check_login(root,cap,Wscreen,light,mat)).grid(row=6,column=2,columnspan=2)

    def check_login(self,root,cap,Wscreen,light,mat):

        if self.the_user.get() == self.credentials['Username']:
            if self.the_pass.get() == self.credentials['Password']:
                self.root.destroy()
                Settings(root,cap,Wscreen,light,mat)
            else:
                self.bad_pass.config(text="Password does not match")
                self.bad_pass.grid(row=2, column=2, columnspan=2)
                self.LogEntry['Password'].delete(0,'end')
        else:
            self.bad_pass.config(text="Username does not exist")
            self.bad_pass.grid(row=2, column=2, columnspan=2)
            self.LogEntry['Username'].delete(0,'end')
            self.LogEntry['Username'].focus()
            self.LogEntry['Password'].delete(0,'end')
    
    # def forget(self):