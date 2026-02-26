import os
import cv2
from tkinter import *
from datetime import datetime
from PIL import Image as PilImg, ImageTk

from Utils.readSettings import readSettings
from Utils.imgProcess import Process

class Accuracy(readSettings):
    def __init__(self,root,image,acc,Wscreen,Hscreen):
        super().__init__()
        self.root = Toplevel(root)
        self.initialize()
        # try:
        self.process(image,Wscreen,Hscreen,acc)
        self.win_config(Wscreen,Hscreen)
        self.widgets(acc,Wscreen,Hscreen)
        self.update()
        self.root.grab_set()
        # except:
        #     self.root.destroy()
        
    def initialize(self):
        self.setAcc, self.AccVar, self.Canva = {},{},{}
        # self.chipArea = float(self.chipSize["02"]['L'])*float(self.chipSize["02"]['W'])

    def win_config(self,Wscreen,Hscreen):
        self.root.title("Accuracy Check Window")
        self.root.state('zoomed')
        self.frame = Frame(self.root)
        self.frame.pack(fill=BOTH, expand=True)

    def widgets(self,acc,Wscreen,Hscreen):
        # Image captured to be displayed in this container
        self.imgWin = Label(self.frame, relief=SUNKEN, image=self.imgtk)
        self.imgWin.grid(row=0, column=0, rowspan=19, columnspan=2, padx=5, pady=10, sticky=NS+EW)

        AccFrame = LabelFrame(self.frame,bd=5,relief=FLAT)
        AccFrame.grid(row=0, column=3, columnspan=2, ipadx=5, padx=5, pady=7, sticky=E)

        ButFrame = LabelFrame(self.frame,bd=5,relief=FLAT)
        ButFrame.grid(row=18, column=3, columnspan=2, ipadx=5, padx=5, pady=7, sticky=E)

        i=0
        for col in self.accuracy[acc]:
            for state in self.accuracy[acc][col]:
                setName = col+state
                self.setAcc[setName] = StringVar(AccFrame, value=self.accuracy[acc][col][state])
                Label(AccFrame, font=self.font['M'], text=setName).grid(row=i, column=0, sticky=W, padx=10, pady=5)
                self.AccVar[setName] = Label(AccFrame, text="0", width=10, relief=RIDGE)
                self.AccVar[setName].grid(row=i, column=1, pady=5, sticky=W)
                Entry(AccFrame, font=self.font['M'], textvariable=self.setAcc[setName], width=12, justify=CENTER).grid(row=i, column=2, sticky=W, padx=10, pady=5)
                self.Canva[setName] = Canvas(AccFrame,width=int(1/40*Wscreen),height=int(1/24*Hscreen))
                self.Canva[setName].grid(row=i, column=3, pady=5, sticky=W)
                self.Canva[setName].create_oval(int(1/320*Wscreen),int(1/192*Hscreen),int(7/320*Wscreen),int(7/192*Hscreen))
                self.Canva[setName].addtag_withtag("circle",1)
                i+=1

        btn_Close = Button(ButFrame, text="Done", font=self.font['M'], height=3, width=23, command=lambda: self.root.destroy())
        btn_Close.grid(row = 0, column = 0, sticky = E)

    def saveImg(self,img,loc,timestp):
        if not self.config['Trouble']:
            imgdir = os.path.join(self.basePath,loc,datetime.today().strftime("%b%y"))
            if not os.path.exists(imgdir): os.makedirs(imgdir)
            imgfile = os.path.join(imgdir,timestp+".png")
            cv2.imwrite(imgfile,img)

    def process(self,imgArr,Wscreen,Hscreen,acc):
        #for ERA03
        chip = "02" if acc == "EQA02" or acc == "EQK02" else "03" if acc =="DMA03" or acc == "ERA03" or acc == "EQK03" or acc == "EQA03" else "15"
        baseimg, image, self.Defects = Process(self.root,imgArr,True,Wscreen,Hscreen,chip,acc=acc).res
        self.saveImg(baseimg,"acc",acc+"_"+datetime.today().strftime("%d-%m-%y"))
        img = PilImg.fromarray(cv2.cvtColor(image,cv2.COLOR_BGR2RGB))
        img = img.resize((int(img.size[0]*0.65),int(img.size[1]*0.65)))
        self.imgtk = ImageTk.PhotoImage(image = img)
        # self.saveImg(baseimg,"test",datetime.today().strftime("%d-%m-%y_%H%M%S"))
        
    def update(self):
        sticker,tape = int(self.tolerance['Sticker']), int(self.tolerance['Tape'])
        for accApp in self.AccVar:
            self.AccVar[accApp].config(text=self.Defects[accApp])
            accCheck = int(self.setAcc[accApp].get())
            if accApp[-4:] == "Area":
                if "Tape" in accApp: passfail = "green" if accCheck - tape < int(self.Defects[accApp]) < accCheck + tape else "red"
                else: passfail = "green" if accCheck - sticker < int(self.Defects[accApp]) < accCheck + sticker else "red"
            else: passfail = "green" if accCheck == int(self.Defects[accApp]) else "red"

            self.Canva[accApp].itemconfig("circle",fill=passfail)