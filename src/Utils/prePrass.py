import os
import pandas as pd
from datetime import datetime

from .readSettings import readSettings
from Pages.Final import showFinal

class prePRASS(readSettings):

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
        self.initialize(root,Wscreen,Hscreen,inData)
        # self.win_config()
        # self.widgets()
        self.Sum_Data(filePath)
        
    def initialize(self,root,Wscreen,Hscreen,inData):
        self.root = root
        self.Wscreen,self.Hscreen = Wscreen,Hscreen
        self.path = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))),"prass",datetime.today().strftime("%b%y"))
        if not os.path.exists(self.path): os.makedirs(self.path)
        self.lotNum,self.mcNo,self.payRoll,self.inQty,self.blkWeight = inData

    # def win_config(self,Wscreen,Hscreen):

    # def widgets(self,Wscreen,Hscreen):

    # Data Cleanup and Summation of Data
    ####################################################################################################
    def Sum_Data(self,filePath):

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

        with open(os.path.join(self.path,f'{self.lotNum}.txt'), 'w') as f: f.write(pData)

        self.showFinal(newdf)

    def showFinal(self,newdf):
        if not self.config['Trouble']: self.res = showFinal(self.root,newdf,self.Wscreen,self.Hscreen,self.lotNum).res
        else: self.res = showFinal(self.root,newdf,self.Wscreen,self.Hscreen).res