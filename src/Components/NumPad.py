from tkinter import *

from Utils.readSettings import readSettings

class Numpad(readSettings):
    
    """
    Tkinter Component with One Entry Box
    Parameters
    ----------
    root : root
        The base root.
    entry : Tkinter Component
        Reference to Entry Box
    Wscreen : str
        Providing Screen Size (Width) to Class.
    Hscreen : str
        Providing Screen Size (Height) to Class.
    """

    def __init__(self,root,entry,Wscreen,Hscreen):
        super().__init__()
        self.root = Toplevel(root)
        # self.initialize()
        self.win_config(Wscreen,Hscreen)
        self.widgets(entry)
        # self.root.grab_set()

    # def initialize(self):

    def win_config(self,Wscreen,Hscreen):
        self.root.geometry(f"{int(Wscreen*0.17)}x{int(Hscreen*0.4)}+{int(Wscreen*0.75)}+20")
        self.WBut = int(Wscreen * 0.005)
        self.HBut = int(Hscreen * 0.005)

    def widgets(self,entry):
        # Create Buttons 1-9 in a 3x3 Mat
        j=0
        for i in range(9):
            if i%3 == 0 and i !=0:
                j+=1
            k = i%3
            Button(self.root,text=i+1,width=self.WBut,height=self.HBut,command=lambda p = i+1: entry.insert(END,p)).grid(row=j,column=k,padx=10,pady=7)
    
        # Create "0" Button
        Button(self.root,text=0,width=self.WBut,height=self.HBut,command=lambda: entry.insert(END,0)).grid(row=4,column=0,pady=5)
        # Create "." Button
        Button(self.root,text=".",width=self.WBut,height=self.HBut,command=lambda: entry.insert(END,".")).grid(row=4,column=1,pady=5)
        # Create "Del" Button
        Button(self.root,text="Del",width=self.WBut,height=self.HBut,command=lambda: entry.delete(entry.index("end") - 1)).grid(row=4,column=2,pady=5)

    