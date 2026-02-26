import os
import requests
import pandas as pd
from datetime import datetime
from tkinter import messagebox

from .readSettings import readSettings
from Pages.Final import showFinal

class PRASS(readSettings):

    """
    Open Lot File, extract and clean data, summarise and send to PRASS.
    Parameters
    ----------
    lotnum : str
        Lot Number.
    mcno : str
        Machine Number.
    payroll : str
        Payroll Number.
    inQty : str
        Input Quantity.
    blkWeight : str
        Block Weight.
    Wscreen : str
        Providing Screen Size (Width) to Class.
    Hscreen : str
        Providing Screen Size (Height) to Class.
    """

    def __init__(self,root,inData,filePath,Wscreen,Hscreen):
        super().__init__()
        self.initialize(root,inData,Wscreen,Hscreen)
        # self.win_config()
        # self.widgets()
        self.Sum_Data(filePath)
        # self.root.grab_set()
        
    def initialize(self,root,inData,Wscreen,Hscreen):
        self.path = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))),"prass",datetime.today().strftime("%b%y"))
        self.root = root
        self.lotNum,self.mcNo,self.payRoll,self.inQty,self.blkWeight = inData
        self.Wscreen,self.Hscreen=Wscreen,Hscreen

    # def win_config(self):

    # def widgets(self):

    # Data Cleanup and Summation of Data
    ####################################################################################################
    def Sum_Data(self,filePath):
        prasspath = os.path.join(self.prassPath,os.path.split(os.path.dirname(filePath))[-1])
        if not os.path.exists(prasspath): os.makedirs(prasspath)

        df = pd.read_excel(filePath, header=None)
        df = df[:].reset_index(drop=True).set_index(df.columns[0])

        df['Total'] = df.sum(axis = 1, skipna = True)
        df.loc['Total Defects'] = df.sum(axis = 0, skipna = True) 
        df = df.iloc[:,-1:]  

        self.create_PRASS(df)

    # Create file spec format for sending to PRASS
    ####################################################################################################
    def create_PRASS(self,df):
        newdf = df.drop(columns=df.columns[0:-1])
        newdf.loc["Output"] = int(self.inQty) - int(newdf.loc['Total Defects','Total'])

        date = datetime.now().strftime("%Y/%m/%d")
        time = datetime.now().strftime("%H:%M:%S")

        df = df.loc[(df!=0).any(axis=1)]

        # Consolidate retrieved defects
        prassDef = ""
        for i, dataQty in enumerate(df["Total"]):
            if i < len(df["Total"]) - 1: prassDef += f"{self.defCode[df.index[i]]}|{dataQty}|" 

        # Range predetermined (Can be changed to increase)
        for j in range(10-(len(df["Total"])-1)):
            if j != range(10-(len(df["Total"])-1))[-1]: prassDef += "||" 
            else: prassDef += "|"

        # File Spec to follow - Added block weight at the end
        pData = f"{self.lotNum}|{self.mcNo}|{self.payRoll}|{date}|{time}|{self.inQty}|{newdf.loc['Output','Total']}|{prassDef}|{self.blkWeight}"

        self.send_PRASS(pData,newdf)

    # Sending txt file to PRASS
    ####################################################################################################
    def send_PRASS(self,pData,newdf):
        self.res = messagebox.askyesno(title="Send Data To PRASS", message="Confirm Send Data?")
        if self.res:
            if not self.config['Trouble']:
                try:
                    with open(os.path.join(self.path,f'{self.lotNum}.txt'), 'w') as f:
                        f.write(pData)

                    # To send via FTP
                    """
                    with FTP("163.50.33.28","s8pc","prasspc") as ftp:
                    ftp.cwd('/fa/thermal_release/')
                    file = open(os.path.join(self.path,f'{self.lotNumb}.txt'),'rb')
                    ftp.storbinary(f'STOR {self.lotNumb}.txt',file)
                    file.close()
                    ftp.quit()
                    """

                    # To Send via HTTP
                    files = {'file': open(os.path.join(self.path,f'{self.lotNum}.txt'), 'rb')}
                    resp = requests.post(self.address["fileServer"], files=files)

                    print(f"fileSize: {int(resp.content)}")
                    # Request OK, bring up website and final quantity for final cross check
                    # if resp.ok:
                    if int(resp.content) == os.stat(os.path.join(self.path,f'{self.lotNum}.txt')).st_size:
                        self.res = showFinal(self.root,newdf,self.Wscreen,self.Hscreen,self.lotNum).res
                except Exception as e:
                    messagebox.showwarning(title="Error while sending",message="File not sent to PRASS. Contact ME.") 
                    print(e)
            else:
                print("Sending to PRASS")
                self.res = showFinal(self.root,newdf,self.Wscreen,self.Hscreen).res