from tkinter import *

from Utils.readSettings import readSettings

class showFinal(readSettings):
    def __init__(self,root,df,Wscreen,Hscreen,lotNum="22X0282300"):
        super().__init__()
        self.root = Toplevel(root)
        self.initialize(Wscreen,Hscreen)
        self.win_config()
        self.widgets(df,lotNum)
        self.root.grab_set()
        self.root.mainloop()

    def initialize(self,Wscreen,Hscreen):
        self.Wscreen, self.Hscreen = Wscreen, Hscreen

    def win_config(self):
        self.root.title("Final Window")
        self.root.geometry(f"{int(self.Wscreen*2/5)}x{int(self.Hscreen*2/3)}+{int(self.Wscreen*3/10)}+{int(self.Hscreen*1/6)}")
        self.frame = Frame(self.root)
        self.frame.columnconfigure(0,weight=1)
        self.frame.rowconfigure(1,weight=1)
        self.frame.pack(fill=BOTH, expand=True)

    def widgets(self,df,lotNum):

        # Labelframe for Lot Number Data
        ####################################################################################################
        lotNumCont = LabelFrame(self.frame, bd=5, relief=FLAT)
        lotNumCont.grid(row=0, column=0, padx=3, pady=3, sticky=NSEW)
        
        lotNumTxt = Label(lotNumCont, text="Lot Number: ", font=self.font['XL'])
        lotNumTxt.grid(row=0, column=0, padx=3, pady=3, sticky=W)

        lotNumb = Label(lotNumCont, text=lotNum, font=self.font['XL'])
        lotNumb.grid(row=0, column=1, padx=3, pady=3)

        # Labelframe for Button
        ####################################################################################################    
        self.butFrame = LabelFrame(self.frame, bd=5, relief=FLAT)
        self.butFrame.grid(row=0, column=1, padx=3, pady=3, sticky=EW)
        
        Button(self.butFrame,text="Okay",bg="#93D976",font=self.font['L'],width=10,pady=5,command=lambda: self.close()).grid(row=0,column=0,sticky=EW)

        # Labelframe for Defects Data
        ####################################################################################################
        self.dataCont = LabelFrame(self.frame, bd=5, relief=FLAT)
        self.dataCont.columnconfigure(2,weight=1)
        self.dataCont.grid(row=1, column=0, columnspan=2, padx=3, pady=3, sticky=NSEW)
        k,l = 0,0
        for i,j in enumerate(df.index[:-1]):
            Label(self.dataCont, text=j,font=self.font['L']).grid(row=l,column=k, pady=3, padx=15, sticky=W)
            Label(self.dataCont, text=df.loc[j][0],font=self.font['L']).grid(row=l,column=k+1, pady=3, padx=15, sticky=EW)
        
            if i%2==0: k=3
            else:
                k = 0 
                l += 1

    def close(self):
        self.res = False
        self.root.destroy()
        self.root.quit()