import os
import cv2
import requests
from tkinter import *
from datetime import datetime
from tkinter import messagebox
from tkinter import Checkbutton, BooleanVar
from PIL import Image as PilImg, ImageTk

from Components.Login import Login
from Components.InputBox import InputBox
from Components.Lighting import Lighting
from Pages.Accuracy import Accuracy
from Pages.Summary import Summary
from Utils.saveExcel import Excel
from Utils.readSettings import readSettings
from Utils.imgProcess import Process
from version import get_full_version


class MainWindow(readSettings):
    def __init__(self, root):
        super().__init__()
        self.root = root
        self.light = Lighting()
        if not self.config["Trouble"]:
            # self.light.initialize()
            self.defVar = {}
            self.defButtons = {}  # Store buttons for manual input mode
            self.defLabels = {}  # Store labels for normal mode
            self.defGridPos = {}  # Store grid positions for defects
            self.evalLotMode = BooleanVar(value=False)  # Evaluation lot toggle state
            self.camera()
            self.win_config()
            self.widgets()

    def camera(self):
        self.cap = cv2.VideoCapture(0, cv2.CAP_DSHOW)
        self.cap.set(cv2.CAP_PROP_FRAME_WIDTH, int(self.config["CamResWidth"]))
        self.cap.set(cv2.CAP_PROP_FRAME_HEIGHT, int(self.config["CamResHeight"]))

    def win_config(self):
        self.root.state("zoomed")
        self.root.title(f"Block Cutting Automation - {get_full_version()}")
        self.Hscreen = self.root.winfo_screenheight()
        self.Wscreen = self.root.winfo_screenwidth()
        self.frame = Frame(self.root)
        self.frame.columnconfigure(0, weight=1)
        self.frame.rowconfigure(len(self.defCode), weight=1)
        self.frame.pack(fill=BOTH, expand=True)
        self.reg = (self.root.register(self.callback), "%P", "%W")

    def widgets(self):
        # Image captured to be displayed in this container
        self.capture = Label(self.frame, relief=SUNKEN)
        self.capture.grid(
            row=0,
            column=0,
            rowspan=len(self.defCode) + 5,
            columnspan=2,
            padx=5,
            pady=10,
            sticky=NS + EW,
        )

        # Frame to hold the defect modes containers
        ####################################################################################################
        defectsInfo = LabelFrame(self.frame, bd=5, relief=FLAT)
        defectsInfo.grid(row=0, column=2, columnspan=2, pady=(7, 2), sticky=EW)

        # Create Label for Defect Code, Defect Name, Defect Quantity based from Naming.py
        k, l = 0, 0
        for j, defName in enumerate(self.defCode):
            Label(defectsInfo, text=f"[{self.defCode[defName]}]").grid(
                row=l, column=k, padx=7, pady=5, sticky=W
            )
            
            # Store grid position for this defect
            self.defGridPos[defName] = {'row': l, 'column': k + 1}
            
            # For DROPCHIP and SAMPLE, always use button
            if defName == "DROPCHIP" or defName == "SAMPLE":
                btn = Button(
                    defectsInfo,
                    text=defName,
                    command=lambda defName=defName: self.SamDropInput(defName),
                )
                btn.grid(row=l, column=k + 1, pady=5, sticky=W)
                self.defButtons[defName] = btn
            else:
                # For other defects, create both label and button (button hidden initially)
                lbl = Label(defectsInfo, text=defName)
                lbl.grid(row=l, column=k + 1, pady=5, sticky=W)
                self.defLabels[defName] = lbl
                
                btn = Button(
                    defectsInfo,
                    text=defName,
                    command=lambda defName=defName: self.SamDropInput(defName),
                )
                # Button initially hidden, will be shown in evaluation mode
                self.defButtons[defName] = btn
            
            self.defVar[defName] = Label(
                defectsInfo, text="0", width=10, relief=RIDGE
            )
            self.defVar[defName].grid(row=l, column=k + 2, pady=5, sticky=W)

            if j % 2 == 0:
                k = 3
            else:
                k = 0
                l += 1

        # Lot Number Scan In
        ####################################################################################################
        lotMcFrame = LabelFrame(self.frame, bd=3, relief=FLAT)
        lotMcFrame.grid(row=len(self.defCode) + 1, column=2, columnspan=2, pady=(0, 2), sticky=EW)
        lotMcFrame.columnconfigure(3, weight=1)

        entrytxt = [
            "Lot Number",
            "M/C Number",
            "PayRoll Number",
            "Blk Weight",
            "Input Quantity",
        ]

        for i, entxt in enumerate(entrytxt):
            # For the first 4 items (indices 0-3), place as before
            if i < 4:
                row_idx = i % 2
                col_idx = 0 if i < 2 else 4
                Label(lotMcFrame, text=entxt, wraplength=50, justify=LEFT).grid(
                    row=row_idx, column=col_idx, pady=5, padx=3, sticky=W
                )
            else:
                # For Input Quantity (index 4), place in row 2, column 0 (directly below M/C Number)
                Label(lotMcFrame, text=entxt, wraplength=50, justify=LEFT).grid(
                    row=2, column=0, pady=5, padx=3, sticky=W
                )

        self.lotNumberEdit = Entry(
            lotMcFrame,
            name="lotno",
            font=self.font["L"],
            width=11,
            justify=CENTER,
            validate="key",
            validatecommand=self.reg,
        )
        self.lotNumberEdit.grid(row=0, column=1, columnspan=2, pady=3, sticky=E)

        self.payRollEdit = Entry(
            lotMcFrame,
            name="payroll",
            font=self.font["L"],
            width=11,
            justify=CENTER,
            validate="key",
            validatecommand=self.reg,
        )
        self.payRollEdit.grid(row=0, column=5, columnspan=2, pady=3, sticky=E)

        self.mcNumberEdit = Entry(
            lotMcFrame,
            name="mcno",
            font=self.font["L"],
            width=11,
            justify=CENTER,
            validate="key",
            validatecommand=self.reg,
        )
        self.mcNumberEdit.grid(row=1, column=1, columnspan=2, pady=3, sticky=E)

        # blkWeightEdit
        blkWeightFrame = Frame(lotMcFrame)
        blkWeightFrame.grid(row=1, column=5, columnspan=2, pady=3, sticky=E)

        self.blkWeightEdit = Entry(
            blkWeightFrame,
            name="blkweight",
            font=self.font["L"],
            width=8,
            justify=CENTER,
            state=DISABLED,
        )
        self.blkWeightEdit.grid(row=0, column=0, pady=0, sticky=W)

        Button(
            blkWeightFrame,
            text="...",
            height=1,
            width=2,
            command=lambda: self.blkWeightInput(),
        ).grid(row=0, column=1, padx=(2, 0), pady=0, sticky=E)

        self.inQtyEdit = Entry(
            lotMcFrame,
            name="inqty",
            font=self.font["L"],
            width=11,
            justify=CENTER,
            validate="key",
            validatecommand=self.reg,
        )
        self.inQtyEdit.grid(row=2, column=1, columnspan=2, pady=3, sticky=E)

        # Evaluation Lot Toggle - placed beside Input Quantity
        ####################################################################################################
        evalCheck = Checkbutton(
            lotMcFrame,
            text="Eval Mode",
            variable=self.evalLotMode,
            command=self.toggleEvalLotMode,
            font=self.font["L"]
        )
        evalCheck.grid(row=2, column=4, padx=(10, 0), pady=3, sticky=W)

        # Spinner for Item Drop Down (02, 03 or 15)
        ####################################################################################################
        AccNChipFrame = LabelFrame(self.frame, bd=3, relief=FLAT)
        AccNChipFrame.grid(row=len(self.defCode) + 3, column=2, columnspan=2, pady=(0, 2), sticky=EW)
        AccNChipFrame.columnconfigure(3, weight=1)

        self.chipType = Label(
            AccNChipFrame, text="ChipType", font=self.font["L"], bg="#ecedcc"
        )
        self.chipType.grid(row=0, column=1, columnspan=3, padx=(0, 10), pady=5)

        accSel = StringVar(value=list(self.accuracy.keys())[0])
        self.accdrop = OptionMenu(AccNChipFrame, accSel, *self.accuracy)
        self.accdrop.config(
            width=int(self.Wscreen * 0.015),
            height=int(self.Hscreen * 0.0027),
            font=self.font["M"],
            anchor=CENTER,
            bg="#c6e2e9",
        )
        self.accdrop.grid(row=0, column=4, columnspan=3, padx=(0, 10), pady=5)

        accSel.trace("w", self.colorCB)

        # Buttons
        ####################################################################################################
        buttonFrame = LabelFrame(self.frame, bd=3, relief=FLAT)
        buttonFrame.grid(row=len(self.defCode) + 4, column=2, columnspan=6, pady=(0, 7), sticky=EW)

        summary = Button(
            buttonFrame,
            text="Summary",
            height=2,
            width=15,
            command=lambda: self.showSum(True),
        )
        summary.grid(row=0, column=0, padx=5, pady=3, sticky=S)

        accuracy = Button(
            buttonFrame,
            text="Accuracy",
            height=2,
            width=15,
            command=lambda: Accuracy(
                self.root, self.prepCam(), accSel.get(), self.Wscreen, self.Hscreen
            ),
        )
        accuracy.grid(row=0, column=1, padx=5, pady=3, sticky=S)

        settings = Button(
            buttonFrame,
            text="Settings",
            height=2,
            width=15,
            command=lambda: Login(
                self.root,
                self.cap,
                self.Wscreen,
                self.Hscreen,
                self.light,
                accSel.get(),
            ),
        )
        settings.grid(row=0, column=2, padx=5, pady=3, sticky=S)

        snap = Button(
            buttonFrame,
            text="Snap",
            height=2,
            width=35,
            font="sans 15 bold",
            bg="#ecedcc",
            command=lambda: self.processImg(
                self.chipType.cget("text")[-2:], accSel.get()
            ),
        )
        snap.grid(row=1, column=0, columnspan=3, padx=10, pady=10, sticky=EW + S)

    ####################################################################################################
    """
    Functions Used for the Widgets Above
    """
    ####################################################################################################

    def blkWeightInput(self):
        """
        Open NumPad for Blk Weight input similar to SamDropInput
        """
        if self.chkEntry(False):
            while True:
                inputValue = InputBox(
                    self.root, "Blk Weight", self.Wscreen, self.Hscreen
                ).inputVal.get()
                if inputValue == "":
                    inputValue = "0"

                # Validate that the input has at most 3 decimal places
                try:
                    float_val = float(inputValue)
                    decimal_part = inputValue.split(".")
                    if len(decimal_part) > 1 and len(decimal_part[1]) > 3:
                        # Show alert if more than 3 decimal places and ask to input again
                        messagebox.showwarning(
                            "Invalid Decimal Places",
                            "Block Weight up to 3 decimal places.\n"
                            f"Your input: {inputValue}\n"
                            "Please re-input again."
                        )
                        # Continue the loop to ask for input again
                        continue
                    else:
                        # Valid input, break out of the loop
                        break
                except ValueError:
                    inputValue = "0"
                    break

            self.blkWeightEdit.config(state=NORMAL)
            self.blkWeightEdit.delete(0, END)
            self.blkWeightEdit.insert(0, inputValue)
            self.blkWeightEdit.config(state=DISABLED)

    def validate_decimal(self, value):
        """
        Validates the Blk Weight input to ensure it has at most 3 decimal places

        Parameters
        ----------
        value : str
            The current value in the entry field

        Returns
        -------
        bool
            True if the input is valid, False otherwise
        """
        # Allow empty input
        if value == "":
            return True

        # Check if the value is a valid decimal number
        try:
            # Try to convert to float to check if it's a valid number
            float_val = float(value)

            # Check if it has at most 3 decimal places
            decimal_part = value.split(".")
            if len(decimal_part) > 1 and len(decimal_part[1]) > 3:
                return False

            return True
        except ValueError:
            # Not a valid float
            return False

    def colorCB(self, name, index, mode):
        if self.root.getvar(name) == "EQA02" or self.root.getvar(name) == "EQK02":
            self.accdrop.config(bg="#c6e2e9")
        elif self.root.getvar(name) == "DMA03" or self.root.getvar(name) == "EQK03" or self.root.getvar(name) == "EQA03":
            self.accdrop.config(bg="#fffdaf")
        else:
            self.accdrop.config(bg="#c7ceea")

    # Checking Entry Boxes and Focus when criteria matches
    ####################################################################################################
    def callback(self, input, name):
        if name.split(".")[-1] == "lotno":
            if len(input) == 10:
                self.inputRetrieve(input)
                foldir = os.path.join(self.dataPath, datetime.today().strftime("%b%y"))
                if not os.path.exists(foldir):
                    os.makedirs(foldir)
                self.filePath = os.path.join(foldir, input + ".xlsx")
                self.payRollEdit.focus()
        elif name.split(".")[-1] == "payroll" and len(input) == 7:
            self.mcNumberEdit.focus()
        elif name.split(".")[-1] == "mcno" and len(input) == 3:
            self.inQtyEdit.focus()  # Skip blkWeightEdit as it's now a button-triggered input
        return True

    # Retrieve Lot number from entry and retrieve Input Quantity
    ####################################################################################################
    def inputRetrieve(self, lotNum):
        url = self.address["QtyWeb"] + lotNum
        res = requests.get(url)
        if res.status_code == 200:
            self.inQtyEdit.delete(0, END)
            self.inQtyEdit.insert(0, res.json()["sun0011"])
            chipCode = (
                ["Error!", "#fa6464"]
                if res.json()["cdc0163"] == None
                else [res.json()["cdc0163"][:5], "#a6eca8"]
            )
            self.chipType.config(text=chipCode[0], bg=chipCode[1])
        else:
            messagebox.showerror(
                "Lot not Found in Database", "Lot not Found in Database"
            )

    # Check for correct input before proceeding to processing image captured
    ####################################################################################################
    def chkEntry(self, var):
        errCode = {
            "lotError": {
                "title": "Lot Number Not Found",
                "message": "Please Input the Lot No",
            },
            "payError": {
                "title": "Payroll Not Found",
                "message": "Please Input the Payroll No",
            },
            "mcError": {
                "title": "Machine Number Not Found",
                "message": "Please Input the Machine No",
            },
            "inputError": {
                "title": "Input Quantity Not Found",
                "message": "Please Input the Input Quantity",
            },
            "blkError": {
                "title": "Block Weight Not Found",
                "message": "Please Input the Block Weight",
            },
        }
        if len(self.lotNumberEdit.get()) != 10:
            return messagebox.showerror(**errCode["lotError"]) != "ok"
        if var:
            if len(self.payRollEdit.get()) == 0:
                return messagebox.showerror(**errCode["payError"]) != "ok"
            elif len(self.mcNumberEdit.get()) == 0:
                return messagebox.showerror(**errCode["mcError"]) != "ok"
            elif len(self.inQtyEdit.get()) == 0:
                return messagebox.showerror(**errCode["inputError"]) != "ok"
            elif len(self.blkWeightEdit.get()) == 0:
                return messagebox.showerror(**errCode["blkError"]) != "ok"
        return True

    # For Editing Entry Box (E.g Sample and Drop Chip) input by User
    ####################################################################################################
    def SamDropInput(self, text):
        if self.chkEntry(False):
            inputValue = InputBox(
                self.root, text, self.Wscreen, self.Hscreen
            ).inputVal.get()
            if inputValue == "":
                inputValue = 0
            self.defVar[text].config(text=inputValue)
            Excel(self.filePath, self.defVar, text)

    # Toggle Evaluation Lot Mode with Credential Check
    ####################################################################################################
    def toggleEvalLotMode(self):
        if self.evalLotMode.get():
            # User wants to enable evaluation mode - check credentials first
            if not self.checkEvalCredentials():
                # Credentials failed, uncheck the toggle
                self.evalLotMode.set(False)
                return
            # Credentials passed, enable evaluation mode
            self.enableEvalMode()
        else:
            # Disable evaluation mode
            self.disableEvalMode()

    # Check Credentials for Evaluation Lot Mode
    ####################################################################################################
    def checkEvalCredentials(self):
        from tkinter import Toplevel, Entry, Label, Button, StringVar
        
        loginWindow = Toplevel(self.root)
        loginWindow.title("Evaluation Lot Login")
        loginWindow.config(bg='lightgrey', bd=50)
        loginWindow.geometry(f"{int(self.Wscreen/4)}x{int(self.Hscreen/4)}+{int(self.Wscreen*3/8)}+{int(self.Hscreen*3/8)}")
        loginWindow.grab_set()
        
        the_user = StringVar()
        the_pass = StringVar()
        
        LogEntry = {}
        LogEntry['Username'] = Entry(loginWindow, textvariable=the_user)
        LogEntry['Password'] = Entry(loginWindow, textvariable=the_pass, show='*')
        bad_pass = Label(loginWindow, bg='red')
        
        for i, cred in enumerate(self.credentials):
            Label(loginWindow, text=cred + ' :', background='lightgrey').grid(row=i, column=1)
            LogEntry[cred].grid(row=i, column=2, columnspan=2)
        
        result = [False]  # Use list to allow modification in nested function
        
        def check_login():
            if the_user.get() == self.credentials['Username']:
                if the_pass.get() == self.credentials['Password']:
                    result[0] = True
                    loginWindow.destroy()
                else:
                    bad_pass.config(text="Password does not match")
                    bad_pass.grid(row=2, column=2, columnspan=2)
                    LogEntry['Password'].delete(0, 'end')
            else:
                bad_pass.config(text="Username does not exist")
                bad_pass.grid(row=2, column=2, columnspan=2)
                LogEntry['Username'].delete(0, 'end')
                LogEntry['Username'].focus()
                LogEntry['Password'].delete(0, 'end')
        
        Button(loginWindow, text="Login", command=check_login).grid(row=6, column=2, columnspan=2)
        
        # Wait for window to close
        loginWindow.wait_window()
        return result[0]

    # Enable Evaluation Mode - Make all defects manually inputtable
    ####################################################################################################
    def enableEvalMode(self):
        for defName in self.defCode:
            if defName != "DROPCHIP" and defName != "SAMPLE":
                # Hide label, show button
                if defName in self.defLabels:
                    self.defLabels[defName].grid_remove()
                if defName in self.defButtons and defName in self.defGridPos:
                    # Get the grid position from stored position
                    pos = self.defGridPos[defName]
                    self.defButtons[defName].grid(
                        row=pos['row'],
                        column=pos['column'],
                        pady=5,
                        sticky=W
                    )

    # Disable Evaluation Mode - Restore normal mode
    ####################################################################################################
    def disableEvalMode(self):
        for defName in self.defCode:
            if defName != "DROPCHIP" and defName != "SAMPLE":
                # Hide button, show label
                if defName in self.defButtons:
                    self.defButtons[defName].grid_remove()
                if defName in self.defLabels and defName in self.defGridPos:
                    # Get the grid position from stored position
                    pos = self.defGridPos[defName]
                    self.defLabels[defName].grid(
                        row=pos['row'],
                        column=pos['column'],
                        pady=5,
                        sticky=W
                    )

    # Preparing Camera for image to process
    ####################################################################################################
    def prepCam(self):
        imgArr = []
        if self.config["Trouble"]:
            imgArr.extend(["", "", ""])
            imgfile = os.path.join(
                self.troublePath, datetime.today().strftime("%d-%m-%y") + ".png"
            )
            imgArr.append(cv2.imread(imgfile))
        else:
            self.light.lightingOn()
            self.cap.release()  # Takes 0.5 seconds to release
            self.camera()  # Takes 1.5 seconds to start up videocapture
            for i in range(13):  # Takes 0.1 seconds to snap each shot per cycle range
                image = self.cap.read()[1]
                imgArr.append(
                    image[
                        :,
                        int(int(self.config["CamResWidth"]) / 6) : int(
                            int(self.config["CamResWidth"]) / 6 * 5
                        ),
                    ]
                )
            self.light.lightingOff()

        return imgArr[3:]

    # Process image and show to User
    ####################################################################################################
    def processImg(self, chip, mat):
        # Force the correct chip type based on material type
        # This ensures EQA02 always uses "02", DMA03 uses "03", and ERA15 uses "15"
        # regardless of what comes from the API or chip type extraction
        chip = "02" if mat == "EQA02" or mat == "EQK02" else "03" if mat == "DMA03" or mat == "EQK03" or mat == "EQA03" or mat == "ERA03" else "15"

        for defName in self.defCode:
            self.defVar[defName].config(text="0")
        if self.chkEntry(True):
            try:
                imgArr = self.prepCam()
                if not imgArr:  # Check if the image array is empty
                    messagebox.showerror(
                        "Image Error", "No images could be captured or loaded"
                    )
                    return

                # Check if evaluation lot mode is enabled
                if self.evalLotMode.get():
                    # Evaluation mode: Only capture and save image, skip defect detection
                    baseimg = imgArr[0] if imgArr else None
                    if baseimg is not None:
                        self.saveImg(
                            baseimg,
                            "block",
                            self.lotNumberEdit.get()
                            + "_"
                            + datetime.today().strftime("%d-%m-%y_%H%M%S"),
                        )
                        # Display the captured image
                        img = PilImg.fromarray(cv2.cvtColor(baseimg, cv2.COLOR_BGR2RGB))
                        img = img.resize((int(img.size[0] * 0.75), int(img.size[1] * 0.75)))
                        imgtk = ImageTk.PhotoImage(image=img)
                        self.capture.imgtk = imgtk
                        self.capture.config(image=imgtk)
                        
                        # Initialize Excel file with defect structure (all zeros)
                        # This ensures the file is ready for manual defect input
                        Excel(self.filePath, self.defVar)
                        
                        # In evaluation mode, user will manually input defects
                        # Don't process defects, just save the image
                        messagebox.showinfo(
                            "Evaluation Mode",
                            "Image captured. Please manually input defect quantities."
                        )
                else:
                    # Normal mode: Process image and detect defects
                    baseimg, image, Defects = Process(
                        self.root, imgArr, False, self.Wscreen, self.Hscreen, chip, mat
                    ).res
                    self.saveImg(
                        baseimg,
                        "block",
                        self.lotNumberEdit.get()
                        + "_"
                        + datetime.today().strftime("%d-%m-%y_%H%M%S"),
                    )
                    img = PilImg.fromarray(cv2.cvtColor(image, cv2.COLOR_BGR2RGB))
                    img = img.resize((int(img.size[0] * 0.75), int(img.size[1] * 0.75)))
                    imgtk = ImageTk.PhotoImage(image=img)
                    self.capture.imgtk = imgtk
                    self.capture.config(image=imgtk)

                    for defappend in self.defVar:
                        if defappend == "DROPCHIP" or defappend == "SAMPLE":
                            addition = (
                                int(self.defVar[defappend].cget("text"))
                                + Defects[defappend]
                            )
                            self.defVar[defappend].config(text=str(addition))
                        else:
                            self.defVar[defappend].config(text=Defects[defappend])

                    Excel(self.filePath, self.defVar)
                    cv2.destroyAllWindows()
            except Exception as e:
                print(e)
                messagebox.showerror(
                    "Processing Error", f"Error processing image: {str(e)}"
                )
                if not self.config["Trouble"]:
                    self.light.lightingOff()

    # Open Summary Window to show Summary of Processed Images
    ####################################################################################################
    def showSum(self, res):
        if self.chkEntry(False) and res:
            inData = [
                self.lotNumberEdit.get(),
                self.mcNumberEdit.get(),
                self.payRollEdit.get(),
                self.inQtyEdit.get(),
                self.blkWeightEdit.get(),
            ]
            self.showSum(
                Summary(
                    self.root, inData, self.filePath, self.Wscreen, self.Hscreen
                ).res
            )
        else:
            self.reset()

    # Reset User Interface for new / next process
    ####################################################################################################
    def reset(self):
        self.lotNumberEdit.delete(0, END)
        self.mcNumberEdit.delete(0, END)
        self.payRollEdit.delete(0, END)
        self.inQtyEdit.delete(0, END)
        self.blkWeightEdit.config(state=NORMAL)  # Temporarily enable to clear
        self.blkWeightEdit.delete(0, END)
        self.blkWeightEdit.config(state=DISABLED)  # Disable again
        self.lotNumberEdit.focus()
        self.capture.config(image="")
        self.chipType.config(text="ChipType", bg="#ecedcc")
        for defName in self.defCode:
            self.defVar[defName].config(text="0")

    # Saving Images
    ####################################################################################################
    def saveImg(self, img, loc, timestp):
        if not self.config["Trouble"]:
            imgdir = os.path.join(self.basePath, loc, datetime.today().strftime("%b%y"))
            if not os.path.exists(imgdir):
                os.makedirs(imgdir)
            imgfile = os.path.join(imgdir, timestp + ".png")
            if not os.path.exists(imgfile):
                cv2.imwrite(imgfile, img)
