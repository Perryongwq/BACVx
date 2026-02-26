import cv2
import math
import numpy as np
from tkinter import messagebox

from Utils.readSettings import readSettings

class Cali(readSettings):

    """
    Mask image to retrieve far left Black pin contour
    Parameters
    ----------
    image : 3-D Image MAT
        Src input image
    """

    def __init__(self,imgArr,mat):
        super().__init__(mat)
        self._initialize()
        self._main(imgArr)

    def _initialize(self):
        self.Area = (int(self.config["Pin Size"])/2)**2*math.pi

    def _main(self,imgArr):
        # pixArea = self.colorMask(img)
        self.error, self.image, pixArea = self.houghCircle(imgArr)
        
        self.avgPixLen = (self.Area/pixArea)
        print(f'Calibration Pin: {self.avgPixLen:.4f} mm^2/pixel')

    # IDEA 1: Mask Image based on Colour (HSV)
    ####################################################################################################
    def colorMask(self,img):
    
        blur = cv2.GaussianBlur(img, (5, 5), 0)
        hsvimg = cv2.cvtColor(blur, cv2.COLOR_BGR2HSV_FULL)
        Black_LL = np.array([int(x) for x in self.color['Pin']['LL'].split(",")], dtype=np.uint8)
        Black_UL = np.array([int(y) for y in self.color['Pin']['UL'].split(",")], dtype=np.uint8)
        black = cv2.inRange(hsvimg,Black_LL,Black_UL)  

        cnts, _ = cv2.findContours(black.copy(), cv2.RETR_EXTERNAL,cv2.CHAIN_APPROX_SIMPLE)

        # Store found contours in dict and sort them to retrieve most left area
        preSort = {}
        for c in cnts:
            cntArea = cv2.contourArea(c)
            x,y,w,h = cv2.boundingRect(c)
            if cntArea > 300:
                if 50 < x & x < 200:
                    preSort[x] = cntArea
                    
        return sorted(preSort.items())[0][1]

    # IDEA 2: Hough Circle to find circle and check if its black
    ####################################################################################################

    def houghCircle(self,imgArr):
        
        pixArr = []
        for img in imgArr:
        
            img = img[800:950,50:250]
            gray = cv2.cvtColor(img,cv2.COLOR_BGR2GRAY)
            blank = np.zeros(img.shape[:2], np.uint8)
            
            circles = cv2.HoughCircles(gray,cv2.HOUGH_GRADIENT,2,20,param1=20,param2=60,minRadius=3,maxRadius=13)
            # print(circles)
            try:
                for c in circles[0]:
                    cv2.circle(blank,(int(c[0]),int(c[1])),int(c[2]),(255,255,255),-1)

                circleMask = cv2.bitwise_and(img,img,mask=blank)

                cnts, _ = cv2.findContours(blank, cv2.RETR_EXTERNAL,cv2.CHAIN_APPROX_SIMPLE)

                blur = cv2.GaussianBlur(circleMask, (5, 5), 0)
                hsvimg = cv2.cvtColor(blur, cv2.COLOR_BGR2HSV_FULL)
                Black_LL = np.array([int(x) for x in self.color['Pin']['LL'].split(",")], dtype=np.uint8)
                Black_UL = np.array([int(y) for y in self.color['Pin']['UL'].split(",")], dtype=np.uint8)
                black = cv2.inRange(hsvimg,Black_LL,Black_UL)  

                for c in cnts:
                    cntArea = cv2.contourArea(c)
                    # print(cntArea,np.sum(black))
                    if cntArea > 300 and np.sum(black)>10000:
                        pixArr.append(cntArea)
            except: pixArr.append("")
        
        # print(pixArr)
        try: 
            NonEmptyArr = [x for x in pixArr if x!=""]
            pixArea = max(set(NonEmptyArr), key = pixArr.count)
            return False, imgArr[pixArr.index(pixArea)], pixArea
        except: 
            messagebox.showerror("Calibration Error","Imaging Issue. \nPlease Try Again.")
            return True, imgArr[0], 1

        
