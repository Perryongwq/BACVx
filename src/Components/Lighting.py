import sys
import socket
from tkinter import messagebox

from Utils.readSettings import readSettings

class Lighting(readSettings):
        
    """
    Lighting Component to connect and execute the Controller
    Parameters
    ----------
    Based from settings
    """
        
    def __init__(self):
        super().__init__()

    def initialize(self):
        if self.config['Trouble']: pass
        else:
            # Establish Connection with controller 
            self.s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            print("connecting ",self.address['LightIPv4'])
            self.s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            try:
                self.s.connect((self.address['LightIPv4'], int(self.address['LightPORT'])))
                print("connected to ",self.address['LightIPv4'])
                self.lightingOff()
            except (ConnectionRefusedError, TimeoutError):
                messagebox.showerror("Connection Error","Connection to Lighting Controller Error. \nPlease Restart Controller.")
                print("Error Connecting to ",self.address['LightIPv4'])
                sys.exit()

    def lightingOn(self):
        if not self.config['Trouble']:
            lightOn = "@00L11D\r\n"
            lightOnarr = []
            for i in lightOn: lightOnarr.append(ord(i))
            bLightOnarr = bytearray(lightOnarr)
            self.s.send(bLightOnarr)
            print(bLightOnarr,"sent")

    def lightingOff(self):
        if not self.config['Trouble']:
            lightOff = "@00L01C\r\n"
            lightOffarr = []
            for i in lightOff: lightOffarr.append(ord(i))
            bLightOffarr = bytearray(lightOffarr)
            self.s.send(bLightOffarr)
            print(bLightOffarr,"sent")

    # Light Intensity Based from Excel Sheet
    def Checksum(self,ComForm):

        Comarr = []
        for i in ComForm: Comarr.append(ord(i))
        ComCheck = bytearray(Comarr)
        CheckSum = hex(sum(ComCheck))[-2:].upper()

        return CheckSum

    # Standardise the lighting intensity digit
    def intenseUtil(self,check):

        if len(check) == 0 or check == 0: check = "100"   # Default
        elif len(check) == 1: check = "00" + check
        elif len(check) == 2: check = "0" + check

        return check

    # Update new Light Intensity
    def lightIntense(self):
        if not self.config['Trouble']:
            intensity = self.intenseUtil(self.config["Lighting"])
            
            preinten = "@00F"+intensity
            checksum = self.Checksum(preinten)
            Command = preinten + checksum+"\r\n"

            arrComm = []
            for i in Command: arrComm.append(ord(i))
            byteCommand = bytearray(arrComm)
            print(byteCommand)
            self.s.send(byteCommand)
            print("sent")