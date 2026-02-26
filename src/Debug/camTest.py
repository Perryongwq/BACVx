import os
import cv2
import sys
import time
import socket
import numpy as np
import pandas as pd
from datetime import datetime
from tkinter import messagebox

from caliTest import CaliTest

address = {"LightIPv4": "192.168.0.2",
            "LightPORT": "40001",}

color = {
            "LL": "0,1,90",
            "UL": "255,40,115"
        }

srcPath = os.path.dirname(os.path.dirname(__file__))
basePath = os.path.dirname(srcPath)

# s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
# print("connecting ",address['LightIPv4'])
# s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
# try:
#     s.connect((address['LightIPv4'], int(address['LightPORT'])))
#     print("connected to ",address['LightIPv4'])
#     lightingOff()
# except (ConnectionRefusedError, TimeoutError):
#     messagebox.showerror("Connection Error","Connection to Lighting Controller Error. \nPlease Restart Controller.")
#     print("Error Connecting to ",address['LightIPv4'])
#     sys.exit()

def lightingOn():
    lightOn = "@00L11D\r\n"
    lightOnarr = []
    for i in lightOn: lightOnarr.append(ord(i))
    bLightOnarr = bytearray(lightOnarr)
    s.send(bLightOnarr)
    print(bLightOnarr,"sent")

def lightingOff():
    lightOff = "@00L01C\r\n"
    lightOffarr = []
    for i in lightOff: lightOffarr.append(ord(i))
    bLightOffarr = bytearray(lightOffarr)
    print(bLightOffarr)
    s.send(bLightOffarr)
    print("sent")

def saveImg(img,loc,timestp):
    imgdir = os.path.join(basePath,loc,datetime.today().strftime("%b%y"))
    if not os.path.exists(imgdir): os.makedirs(imgdir)
    imgfile = os.path.join(imgdir,timestp+".png")
    if not os.path.exists(imgfile): cv2.imwrite(imgfile,img)

baseimg,pixArea = CaliTest()

saveImg(baseimg,"test",datetime.today().strftime("%d-%m-%y_%H%M%S"))
