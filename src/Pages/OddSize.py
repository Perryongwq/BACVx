import cv2
from tkinter import *
from tkinter import messagebox
from PIL import Image as PilImg, ImageTk

from Utils.readSettings import readSettings

class OddSize(readSettings):
    def __init__(self,root,oddDict,Bimg,Wscreen,Hscreen):
        super().__init__()
        self.root = Toplevel(root)
        self.initialize(Bimg)
        self.win_config(Wscreen,Hscreen)
        self.widgets(oddDict)
        self.root.grab_set()
        self.root.mainloop()
        
    def initialize(self,img):
        img = PilImg.fromarray(cv2.cvtColor(img,cv2.COLOR_BGR2RGB))
        img = img.resize((int(img.size[0]*0.75),int(img.size[1]*0.75)))
        self.imgtk = ImageTk.PhotoImage(image = img)
        self.choose1 = "Choose 1"
        self.frameHold, self.dropbox, self.config, self.selected ={},{},{},{}

    def win_config(self,Wscreen,Hscreen):
        self.root.title("Tape Defect Window")
        self.root.geometry(f"{int(Wscreen*0.7)}x{int(Hscreen*0.9)}+30+10")
        self.frame = Frame(self.root)
        self.frame.columnconfigure(0,weight=1)
        self.frame.rowconfigure(9,weight=1)
        self.frame.pack(fill=BOTH, expand=True)
        
    def widgets(self,oddDict):
        # Image captured to be displayed in this container
        self.imgWin = Label(self.frame, relief=SUNKEN, image=self.imgtk)
        self.imgWin.grid(row=0, column=0, rowspan=19, columnspan=2, padx=5, pady=10, sticky=NS+EW)

        # Button to Save Selection and return Black Sticker Defects
        btn_OddStick = Button(self.frame, text="Save Changes", font=self.font['M'], height=2, width=15, command=lambda: self.saveSelected(oddDict))
        btn_OddStick.grid(row=10, column=len(oddDict)*2, columnspan=2, pady=20, padx=10, sticky=NE)     
        
        # Set up Tape layout but not injected into Tkinter yet
        ##################################################################################################
        i = 1
        for key, value in oddDict.items():
            if key[-4:] in self.highlight.keys():
                self.frameHold[key] = LabelFrame(self.frame,bd=5,relief=FLAT,text=key)
                self.frameHold[key].grid(row=0, column=2*i, columnspan=2, ipadx=5, padx=5, pady=7, sticky=E)
                i+=1
                self.dropbox[key], self.config[key] = {}, {}
                for j in range(len(value)):
                    Label(self.frameHold[key], text=f"[{j+1}]").grid(row=j, column=0, pady=3, padx=10, sticky=W)
                    self.dropbox[key][j] = StringVar(self.frameHold[key])
                    self.dropbox[key][j].set(self.choose1)
                    if "Tape" in key:
                        self.config[key][j] = OptionMenu(self.frameHold[key],self.dropbox[key][j],*self.defCode)
                        self.config[key][j].config(width=13)
                        self.config[key][j].grid(row=j, column=1, pady=3, sticky=W)
                    else:
                        self.config[key][j] = OptionMenu(self.frameHold[key],self.dropbox[key][j],*self.defCode)
                        self.config[key][j].config(width=13)
                        self.config[key][j].grid(row=j, column=1, pady=3, sticky=W)

    def saveSelected(self,oddDict):

        for key, value in oddDict.items():
            if key[-4:] in self.highlight.keys():
                self.selected[key] = {}
                for k in range(len(value)):
                    if self.dropbox[key][k].get() in self.selected[key].keys(): self.selected[key][self.dropbox[key][k].get()] += oddDict[key][k]
                    else: self.selected[key][self.dropbox[key][k].get()] = oddDict[key][k]
            a = 0
        for ky in self.selected.keys():
            if not self.selected[ky] == {}:
                if "Choose 1" in self.selected[ky].keys() and a == 0:
                        messagebox.showwarning("Empty Selection","Please Choose 1!",parent=self.root)
                        a+=1                
                else: 
                    self.root.destroy()
                    self.root.quit()