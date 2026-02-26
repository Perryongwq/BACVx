import os
import cv2
import json
import numpy as np
from tkinter import *
from tkinter import ttk
from datetime import datetime
from tkinter import messagebox
from PIL import Image, ImageTk

from Utils.readSettings import readSettings

class Settings(readSettings):
    def __init__(self,root,cap,Wscreen,light,mat):
        super().__init__(mat)
        self.light = light
        self.root = Toplevel(root)
        self.initialize(cap,Wscreen)
        self.win_config()
        self.widgets(mat)
        light.lightingOn()
        self.root.grab_set()

    def initialize(self,cap,Wscreen):
        self.cap = cap
        self.camW = int(self.config['CamResWidth'])   
        self.length = int(Wscreen*0.008)
        self.font = ('Calibri',self.length)
        self.colTxt = 'Select Colour:'
        self.setTxt = 'Select Setting:'
        self.edit = 'Edit'
        self.HSVtext=['Low H: ', 'High H: ', 'Low S: ', 'High S: ', 'Low V: ', 'High V: ']
        self.hsvLabel,self.hsvVal,self.colLL,self.colUL = {},{},{},{}
        self.chipL,self.chipW,self.accNum,self.accChip  = {},{},{},{}
        self.fig,self.addr,self.cred,self.accSel,self.btn5 = {},{},{},{},{}
        self.addrVar,self.figVar,self.chipLVar,self.chipWVar = {},{},{},{}
        self.credVar,self.accNumVar,self.accChipVar = {},{},{}
        self.tolVar,self.tole,self.tapeVar,self.tape = {},{},{},{}

    def win_config(self):
        self.root.title("Settings")
        self.root.state('zoomed')
        self.root.columnconfigure(1,weight=1)
        self.root.columnconfigure(2,weight=1)
        self.root.columnconfigure(3,weight=1)

    def widgets(self,mat):
        # Frame Creation x 5 (Left: Video Frame, Center Top: TrackBar for Video Frame, Center Bottom: Misc Config, Right: Colour Ranges, Bottom right: Buttons)
        self.camFrame = Frame(self.root,highlightbackground='blue',highlightthickness=2)
        self.camFrame.grid(row=0,column=0,rowspan=2,sticky=NS)
        self.slideFrame = Frame(self.root)
        self.slideFrame.grid(row=0,column=2,padx=10,sticky=NSEW)
        self.editFrame = Frame(self.root,bg='#999',highlightbackground='blue',highlightthickness=2)
        self.editFrame.grid(row=1,column=2,sticky=NSEW)
        self.editFrame.rowconfigure(0,weight=1)
        self.editFrame.columnconfigure(0,weight=1)
        self.colourFrame = Frame(self.root,highlightbackground='blue',highlightthickness=2)
        self.colourFrame.grid(row=0,column=4,rowspan=2,padx=10,sticky=NS+EW)
        self.butFrame = Frame(self.root)
        self.butFrame.grid(row=1,column=4,padx=10,sticky=NS)

        # Creation of Sliders / Trackbars
        instructions = 'Instructions:\nMove trackbar to achieve the correct HSV for each colour.\nRecord the range as HSV_Low and HSV_High range.\n\nExample - HSV_Low=0,0,0 and HSV_High=255,255,255\n'
        Label(self.slideFrame,text=instructions,font=self.font,justify=LEFT,anchor=W).grid(row=0,column=0,columnspan=4,sticky=W+E,padx=10,pady=(10,0))
        
        for i, txt in enumerate(self.HSVtext):
            self.hsvVal[txt] = ttk.Scale(self.slideFrame,from_=0,to=255,orient=HORIZONTAL,length=self.length*9, command=lambda event,x=txt,y=i: self.hsvLabel[x].configure(text=self.HSVtext[y]+str(int(float(event)))))
            self.hsvVal[txt].grid(row=i+1,column=1,columnspan=3,sticky=W+E,padx=10,pady=(0,5))
            self.hsvLabel[txt] = Label(self.slideFrame,font=self.font,text=self.HSVtext[i]+str(int(self.hsvVal[txt].get())))
            self.hsvLabel[txt].grid(row=i+1,column=0,sticky=W)
            self.hsvVal[txt].set(0) if i%2 == 0 else self.hsvVal[txt].set(255)

        # Select Dropdown box to choose color
        self.dropColSel = StringVar()
        self.dropColSel.set(self.colTxt)
        coldrop = OptionMenu(self.slideFrame,self.dropColSel,*self.color)
        coldrop.config(width=int(self.length*1.3))
        coldrop.grid(row=7,column=0,padx=(0,10),pady=20)

        # Link Buttons to Fetch, Add data and reset Sliders / Trackbars
        Button(self.slideFrame,text='Fetch',font=self.font,width=self.length,command=lambda: self.fetch()).grid(row=7,column=1,padx=10,pady=20)
        Button(self.slideFrame,text='Add',font=self.font,width=self.length,command=lambda: self.add()).grid(row=7,column=2,padx=10,pady=20)
        Button(self.slideFrame,text='Reset',font=self.font,width=self.length,command=lambda: self.reset()).grid(row=7,column=3,padx=10,pady=20)

        # Creation of Colour Ranges columns
        Label(self.colourFrame,text='LL',font=self.font,width=self.length).grid(row=0,column=1,pady=5)
        Label(self.colourFrame,text='UL',font=self.font,width=self.length).grid(row=0,column=2,pady=5)

        for j, col in enumerate(self.color):
            Label(self.colourFrame,text=col,font=self.font,width=self.length).grid(row=j+1,column=0,padx=5,pady=5)
            self.colLL[col] = Label(self.colourFrame,font=self.font,justify=CENTER,fg='#777',width=self.length,text=self.color[col]['LL'])
            self.colUL[col] = Label(self.colourFrame,font=self.font,justify=CENTER,fg='#777',width=self.length,text=self.color[col]['UL'])
            self.colLL[col].grid(row=j+1,column=1,padx=5,pady=5)
            self.colUL[col].grid(row=j+1,column=2,padx=5,pady=5)
        
        
        # Creation of Misc Configs / Settings TabContainer
        tabCont = ttk.Notebook(self.editFrame)
        tab1 = ttk.Label(tabCont)
        tab2 = ttk.Label(tabCont)
        tab3 = ttk.Label(tabCont)
        tab4 = ttk.Label(tabCont)
        tab5 = ttk.Label(tabCont)
        tab6 = ttk.Label(tabCont)

        tabCont.add(tab1,text='Address')
        tabCont.add(tab2,text='Config')
        tabCont.add(tab3,text='Chip Size')
        tabCont.add(tab4,text='Credentials')
        tabCont.add(tab5,text='Accuracy')
        tabCont.add(tab6,text='Tolerance')
        tabCont.grid(row=0,column=0,padx=15,pady=30,sticky=NSEW)

        # [Tab 1] Address Tab
        ####################################################################################################
        for a, addr in enumerate(self.address):
            ttk.Label(tab1,text=addr,font=self.font,width=int(self.length*1.3)).grid(row=a,column=0,padx=(10,5),pady=5)

            self.addrVar[addr] = StringVar(value=self.address[addr])
            self.addr[addr] = ttk.Entry(tab1,font=self.font,justify=CENTER,state=DISABLED,foreground='#777',textvariable=self.addrVar[addr],width=self.length*4)
            self.addr[addr].grid(row=a,column=1,columnspan=2,padx=5,pady=5,sticky=EW)

        # Select Dropdown box for Misc
        self.confSel = StringVar(value=self.setTxt)
        confdrop = OptionMenu(tab1,self.confSel,*self.addr)
        confdrop.config(width=int(self.length*1.3))
        confdrop.grid(row=a+1,column=1,padx=(0,10),pady=20,sticky=N)

        btn1 = Button(tab1,text=self.edit,font=self.font,width=self.length,command=lambda: self.modify(tabCont.tab(tabCont.select(),'text'),btn1))
        btn1.grid(row=a+1,column=2,padx=10,pady=20,sticky=N)

        # [Tab 2] Config Tab
        ####################################################################################################
        for b, conf in enumerate(self.config):
            ttk.Label(tab2,text=conf,font=self.font,width=int(self.length*1.3)).grid(row=b,column=0,padx=(10,5),pady=5)

            if conf == "Trouble":
                s = ttk.Style()
                s.configure('Radio.TRadiobutton',font=self.font,justify=CENTER,foreground='#777')

                self.figVar[conf] = BooleanVar(value=self.config[conf])
                self.fig[conf+"True"] = ttk.Radiobutton(tab2,text="True",value=1,state=DISABLED,variable=self.figVar[conf],style='Radio.TRadiobutton')
                self.fig[conf+"True"].grid(row=b,column=1,padx=5,pady=5,sticky=EW)
                self.fig[conf+"False"] = ttk.Radiobutton(tab2,text="False",value=0,state=DISABLED,variable=self.figVar[conf],style='Radio.TRadiobutton')
                self.fig[conf+"False"].grid(row=b,column=2,padx=5,pady=5,sticky=EW)
            else:
                self.figVar[conf] = StringVar(value=self.config[conf])
                self.fig[conf] = ttk.Entry(tab2,font=self.font,justify=CENTER,state=DISABLED,foreground='#777',textvariable=self.figVar[conf],width=self.length*4)
                self.fig[conf].grid(row=b,column=1,columnspan=2,padx=5,pady=5,sticky=EW)

        # Select Dropdown box for Misc
        self.figSel = StringVar(value=self.setTxt)
        figdrop = OptionMenu(tab2,self.figSel,*self.config)
        figdrop.config(width=int(self.length*1.3))
        figdrop.grid(row=b+1,column=1,padx=(0,10),pady=20,sticky=N)

        btn2 = Button(tab2,text=self.edit,font=self.font,width=self.length,command=lambda: self.modify(tabCont.tab(tabCont.select(),'text'),btn2))
        btn2.grid(row=b+1,column=2,padx=10,pady=20,sticky=N)

        # [Tab 3] Chips Tab
        ####################################################################################################
        ttk.Label(tab3,text='Length',font=self.font,width=self.length,anchor='center').grid(row=0,column=1,pady=(5,0))
        ttk.Label(tab3,text='Width',font=self.font,width=self.length,anchor='center').grid(row=0,column=2,pady=(5,0))

        for c, chip in enumerate(self.chipSize):
            ttk.Label(tab3,text=chip,font=self.font,width=int(self.length*1.3)).grid(row=c+1,column=0,padx=(10,5),pady=5)

            self.chipLVar[chip] = StringVar(value=self.chipSize[chip]['L'])
            self.chipL[chip] = ttk.Entry(tab3,font=self.font,justify=CENTER,state=DISABLED,foreground='#777',textvariable=self.chipLVar[chip],width=self.length*2)
            self.chipL[chip].grid(row=c+1,column=1,padx=5,pady=5)
            self.chipWVar[chip] = StringVar(value=self.chipSize[chip]['W'])
            self.chipW[chip] = ttk.Entry(tab3,font=self.font,justify=CENTER,state=DISABLED,foreground='#777',textvariable=self.chipWVar[chip],width=self.length*2)
            self.chipW[chip].grid(row=c+1,column=2,padx=5,pady=5)

        # Select Dropdown box for Chips
        self.chipSel = StringVar(value=self.setTxt)
        chipdrop = OptionMenu(tab3,self.chipSel,*self.chipSize)
        chipdrop.config(width=int(self.length*1.3))
        chipdrop.grid(row=c+2,column=1,padx=(0,10),pady=20,sticky=N)

        btn3 = Button(tab3,text=self.edit,font=self.font,width=self.length,command=lambda: self.modify(tabCont.tab(tabCont.select(),'text'),btn3))
        btn3.grid(row=c+2,column=2,padx=10,pady=20,sticky=N)

        # [Tab 4] Credentials Tab
        ####################################################################################################
        for d, cred in enumerate(self.credentials):
            ttk.Label(tab4,text=cred,font=self.font,width=int(self.length*1.3)).grid(row=d,column=0,padx=(10,5),pady=5)

            self.credVar[cred] = StringVar(value=self.credentials[cred])
            self.cred[cred] = ttk.Entry(tab4,font=self.font,justify=CENTER,state=DISABLED,foreground='#777',textvariable=self.credVar[cred],width=self.length*4)
            self.cred[cred].grid(row=d,column=1,columnspan=2,padx=5,pady=5,sticky=EW)

        # Select Dropdown box for Misc
        self.credSel = StringVar(value=self.setTxt)
        creddrop = OptionMenu(tab4,self.credSel,*self.credentials)
        creddrop.config(width=int(self.length*1.3))
        creddrop.grid(row=d+1,column=1,padx=(0,10),pady=20,sticky=N)

        btn4 = Button(tab4,text=self.edit,font=self.font,width=self.length,command=lambda: self.modify(tabCont.tab(tabCont.select(),'text'),btn4))
        btn4.grid(row=d+1,column=2,padx=10,pady=20,sticky=N)

        # [Tab 5] Accuracy Tab
        ####################################################################################################
        tab5cont = ttk.Notebook(tab5)
        innertab5,accdrop = {},{}
        for acc in self.accuracy:
            innertab5[acc] = ttk.Label(tab5cont)
            tab5cont.add(innertab5[acc],text=acc)
            tab5cont.grid(row=0,column=0)

            ttk.Label(innertab5[acc],text='Num',font=self.font,width=self.length,anchor='center').grid(row=0,column=1,pady=(5,0))
            ttk.Label(innertab5[acc],text='Area',font=self.font,width=self.length,anchor='center').grid(row=0,column=2,pady=(5,0))

            for e, accCol in enumerate(self.accuracy[acc]):
                ttk.Label(innertab5[acc],text=accCol,font=self.font,width=self.length).grid(row=e+1,column=0,padx=(10,5),pady=5)

                self.accNumVar[acc+accCol] = StringVar(value=self.accuracy[acc][accCol]['Num'])
                self.accNum[acc+accCol] = ttk.Entry(innertab5[acc],font=self.font,justify=CENTER,state=DISABLED,foreground='#777',textvariable=self.accNumVar[acc+accCol],width=self.length*2)
                self.accNum[acc+accCol].grid(row=e+1,column=1,padx=5,pady=5,sticky=EW)
                self.accChipVar[acc+accCol] = StringVar(value=self.accuracy[acc][accCol]['Area'])
                self.accChip[acc+accCol] = ttk.Entry(innertab5[acc],font=self.font,justify=CENTER,state=DISABLED,foreground='#777',textvariable=self.accChipVar[acc+accCol],width=self.length*2)
                self.accChip[acc+accCol].grid(row=e+1,column=2,padx=5,pady=5,sticky=EW)

            # Select Dropdown box for Misc
            self.accSel[acc] = StringVar(value=self.setTxt)
            accdrop[acc] = OptionMenu(innertab5[acc],self.accSel[acc],*self.accuracy[acc])
            accdrop[acc].config(width=int(self.length*1.3))
            accdrop[acc].grid(row=e+2,column=1,padx=(0,10),pady=20,sticky=N)

            self.btn5[acc] = Button(innertab5[acc],text=self.edit,font=self.font,width=self.length,command=lambda acc=acc: self.modify(tabCont.tab(tabCont.select(),'text'),self.btn5[acc],acc))
            self.btn5[acc].grid(row=e+2,column=2,padx=10,pady=20,sticky=N)

        # [Tab 6] Tolerance Tab
        ####################################################################################################
        for f, tolType in enumerate(self.tolerance):
            ttk.Label(tab6,text=tolType,font=self.font,width=int(self.length*1.3)).grid(row=f+1,column=0,padx=(10,5),pady=5)

            self.tolVar[tolType] = StringVar(value=self.tolerance[tolType])
            self.tole[tolType] = ttk.Entry(tab6,font=self.font,justify=CENTER,state=DISABLED,foreground='#777',textvariable=self.tolVar[tolType],width=self.length*4)
            self.tole[tolType].grid(row=f+1,column=1,columnspan=2,padx=5,pady=5)

        # Select Dropdown box for tolerance per chip
        self.tolSel = StringVar(value=self.setTxt)
        toldrop = OptionMenu(tab6,self.tolSel,*self.tolerance)
        toldrop.config(width=int(self.length*1.3))
        toldrop.grid(row=f+2,column=1,padx=(0,10),pady=20,sticky=N)

        btn6 = Button(tab6,text=self.edit,font=self.font,width=self.length,command=lambda: self.modify(tabCont.tab(tabCont.select(),'text'),btn6))
        btn6.grid(row=f+2,column=2,padx=10,pady=20,sticky=N)

        # Save Changes or No changes
        Button(self.butFrame,text='Save Changes',width=self.length,command=lambda: self.save(mat)).grid(row=0,column=0,padx=10,pady=(10,20),ipadx=10,ipady=5,sticky=E)
        Button(self.butFrame,text='No Changes',width=self.length,command=lambda: self.quit()).grid(row=0,column=1,padx=(10,0),pady=(10,20),ipadx=10,ipady=5,sticky=E)

        # Place image / Video frame into frame
        self.capture = Label(self.camFrame)
        self.capture.grid(row=0,column=0)
        self.show_frame()

    #Display to image frame in Window
    def show_frame(self):
        if self.config['Trouble']: 
            frame = cv2.imread(os.path.join(self.troublePath,datetime.today().strftime("%d-%m-%y")+".png"))
        else:
            _, frame = self.cap.read()
            frame = frame[:,int(self.camW/6):int(self.camW/6*5)]

        hsv = cv2.cvtColor(frame, cv2.COLOR_BGR2HSV_FULL) #convert source image to HSV color mode
        hsv_low = np.array([self.hsvVal['Low H: '].get(), self.hsvVal['Low S: '].get(), self.hsvVal['Low V: '].get()], np.uint8)
        hsv_high = np.array([self.hsvVal['High H: '].get(), self.hsvVal['High S: '].get(), self.hsvVal['High V: '].get()], np.uint8)

        # making mask for hsv range
        mask = cv2.inRange(hsv, hsv_low, hsv_high)
        vis = cv2.cvtColor(mask, cv2.COLOR_GRAY2BGR)

        # masking HSV value selected color becomes black
        rgbframe=cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        res = cv2.bitwise_and(rgbframe, rgbframe, mask=mask)

        h1, w1 = res.shape[:2]
        h2, w2 = vis.shape[:2]
        ##########################VERTICAL############################
        concat=np.zeros((h1+h2, max(w1,w2),3), dtype=np.uint8)
        concat[:,:] = (255,255,255)
        concat[:h1, :w1,:3] = res
        concat[h1:h1+h2, :w2,:3] = vis

        img = Image.fromarray(concat)
        width=int(self.root.winfo_height()/(img.size[1]/img.size[0]))
        height=self.root.winfo_height()
        if width==0 or height==0: width,height=(1,1)
        img=img.resize((width, height))
        imgtk = ImageTk.PhotoImage(image=img)
        self.capture.imgtk = imgtk
        self.capture.configure(image=imgtk)
        self.capture.after(10, self.show_frame)

    # Fetch HSV setting for colour from settings
    def fetch(self):
        colour=self.dropColSel.get()
        if colour!=self.colTxt:
            hsvLL = self.colLL[colour].cget('text').split(',')
            hsvUL = self.colUL[colour].cget('text').split(',')

            for i, txt in enumerate(self.HSVtext):
                self.hsvVal[txt].set(hsvLL[i//2]) if i%2 == 0 else self.hsvVal[txt].set(hsvUL[i//2])                
        else:
            self.msgBox('fetch')

    # Add HSV for colour to settings
    def add(self):
        colour=self.dropColSel.get()
        if colour!=self.colTxt:
            hsvLL, hsvUL = [],[]
            for i, txt in enumerate(self.HSVtext):
                if i%2 == 0:
                    hsvLL.append(str(int(self.hsvVal[txt].get())))
                else:
                    hsvUL.append(str(int(self.hsvVal[txt].get())))

            self.colLL[colour].config(text=','.join(hsvLL))
            self.colUL[colour].config(text=','.join(hsvUL))
            messagebox.showinfo('Prompt', 'Successfully added!',parent=self.root)
        else:
            self.msgBox('add')

    # Reset HSV for all trackbars
    def reset(self):
        for i, txt in enumerate(self.HSVtext):
            self.hsvVal[txt].set(0) if i%2 == 0 else self.hsvVal[txt].set(255)
            self.hsvLabel[txt].configure(text=self.HSVtext[i]+str(int(self.hsvVal[txt].get())))

    def modify(self,tab,btn,var=0):

        if tab == 'Address':
            if self.confSel.get() == self.setTxt: return
        elif tab == 'Config':
            if self.figSel.get() == self.setTxt: return
        elif tab == 'Chip Size':
            if self.chipSel.get() == self.setTxt: return
        elif tab == 'Credentials':
            if self.credSel.get() == self.setTxt: return
        elif tab == 'Accuracy':
            if self.accSel[var].get() == self.setTxt: return
        elif tab == 'Tolerance':
            if self.tolSel.get() == self.setTxt: return

        if btn.cget('text') == 'Edit':
            btn.config(text='Save Changes')
            self.tabSel(tab,NORMAL)
        elif btn.cget('text') == 'Save Changes':
            btn.config(text='Edit')
            self.tabSel(tab,DISABLED)
    
    def tabSel(self,tab,state):
        if tab == 'Address':
            for i in self.address:
                if self.confSel.get() == i:
                    self.addr[i].configure(state=state)
        elif tab == 'Config':
            for i in self.config:
                if self.figSel.get() == i:
                    if self.figSel.get() == "Trouble":
                        self.fig[i+"True"].configure(state=state)
                        self.fig[i+"False"].configure(state=state)
                    else:
                        self.fig[i].configure(state=state)
        elif tab == 'Chip Size':
            for i in self.chipSize:
                if self.chipSel.get() == i:
                    self.chipL[i].configure(state=state)
                    self.chipW[i].configure(state=state)
        elif tab == 'Credentials':
            for i in self.credentials:
                if self.credSel.get() == i:
                    self.cred[i].configure(state=state)
        elif tab == 'Accuracy':
            for i in self.accuracy:
                for j in self.setData[tab][i]:
                    if self.accSel[i].get() == j:
                        self.accNum[i+j].configure(state=state)
                        self.accChip[i+j].configure(state=state)
        elif tab == 'Tolerance':
            for i in self.tolerance:
                if self.tolSel.get() == i:
                    self.tole[i].configure(state=state)
        else:
            pass

    def save(self,mat):
        for col in self.color:
            self.miscData['Color'][col]['LL'] = self.colLL[col].cget('text')
            self.miscData['Color'][col]['UL'] = self.colUL[col].cget('text')
        
        for chip in self.chipSize:
            self.setData['Chip Size'][chip]['L'] = self.chipL[chip].get()
            self.setData['Chip Size'][chip]['W'] = self.chipW[chip].get()

        for conf in self.config:
            self.setData['Config'][conf] = self.figVar[conf].get() if conf == "Trouble" else self.fig[conf].get()

        for cred in self.credentials:
            self.setData['Credentials'][cred] = self.cred[cred].get()

        for acc in self.accuracy:
            for accCol in self.accuracy[acc]:
                self.setData['Accuracy'][acc][accCol]['Num'] = self.accNum[acc+accCol].get()
                self.setData['Accuracy'][acc][accCol]['Area'] = self.accChip[acc+accCol].get()

        for tol in self.tolerance:
            self.setData['Tolerance'][tol] = self.tole[tol].get()

        with open(os.path.join(self.srcPath, 'JSON', 'settings.json'),"w") as f:
            f.seek(0)
            json.dump(self.setData,f,indent=4)
            f.truncate()

        with open(os.path.join(self.srcPath, 'JSON', self.machine, f'{mat}misc.json'),"w") as f:
            f.seek(0)
            json.dump(self.miscData,f,indent=4)
            f.truncate()
        
        messagebox.showinfo('Prompt', 'Changes have been saved!', parent=self.root)

        self.light.lightIntense()
        self.light.lightingOff()
        self.root.destroy()

    def msgBox(self,text):
        messagebox.showinfo('Prompt',f'Select a colour to {text}', parent=self.root)

    def quit(self):
        self.light.lightIntense()
        self.light.lightingOff()
        self.root.destroy()
        