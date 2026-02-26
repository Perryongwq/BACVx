
import os
import cv2
import math
import time
import numpy as np
import pandas as pd

from Utils.readSettings import readSettings

class Evaluation(readSettings):
    def __init__(self,img,chip=False,acc=False):
        super().__init__()
        self.initialize(chip,acc)
        image, self.avgPixLen =self.calibration(img)
        resImg = self.masking(image)
        stickers = self.findSticker(resImg.copy())
        cnts, hier = cv2.findContours(stickers,cv2.RETR_EXTERNAL,cv2.CHAIN_APPROX_SIMPLE)
        self.chipPixelCnt(resImg,cnts)

    def initialize(self,chip,acc):
        self.colDict = {}
        self.chipArea = 1
        for col in self.color.keys():
            if col == "Pin": continue
            if chip:
                self.chipArea = float(self.chipSize[chip]['L'])*float(self.chipSize[chip]['W'])
                if chip in self.color[col]['chipType']:
                    Col_LL = np.array([int(x) for x in self.color[col]['LL'].split(",")], dtype=np.uint8)
                    Col_UL = np.array([int(y) for y in self.color[col]['UL'].split(",")], dtype=np.uint8)
                    self.colDict[col] = {"LL" : Col_LL, "UL" : Col_UL}
            else:
                if col in self.accuracy[acc].keys():
                    Col_LL = np.array([int(x) for x in self.color[col]['LL'].split(",")], dtype=np.uint8)
                    Col_UL = np.array([int(y) for y in self.color[col]['UL'].split(",")], dtype=np.uint8)
                    self.colDict[col] = {"LL" : Col_LL, "UL" : Col_UL}

    def calibration(self,img):
        
        pixArr = []
        baseImg = img.copy()       
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
        try :
            NonEmptyArr = [x for x in pixArr if x!=""]
            pixArea = max(set(NonEmptyArr), key = pixArr.count)

            cali = ((int(self.config["Pin Size"])/2)**2*math.pi/pixArea)
        except: cali = 0.05674842221079829

        return baseImg, cali

    def masking(self,img):

        # Variables
        blank = np.zeros(img.shape[:2], np.uint8)
        kernel = np.ones((5,5), np.uint8)
        ker = np.ones((51,51),np.uint8)

        # Find Block
        blur = cv2.bilateralFilter(img.copy(),9,15,15)

        m1 = cv2.adaptiveThreshold(blur[:,:,0],255,cv2.ADAPTIVE_THRESH_GAUSSIAN_C,cv2.THRESH_BINARY_INV,15,3)
        m2 = cv2.adaptiveThreshold(blur[:,:,1],255,cv2.ADAPTIVE_THRESH_GAUSSIAN_C,cv2.THRESH_BINARY_INV,15,3)
        m3 = cv2.adaptiveThreshold(blur[:,:,2],255,cv2.ADAPTIVE_THRESH_GAUSSIAN_C,cv2.THRESH_BINARY_INV,15,3)
        thresh = cv2.add(m1, cv2.add(m2, m3))

        dilate = cv2.dilate(thresh,kernel)
        cnt, hier = cv2.findContours(dilate,cv2.RETR_EXTERNAL,cv2.CHAIN_APPROX_SIMPLE)
        cn = sorted(cnt, key=lambda x: cv2.contourArea(x))[-1]
        mask = cv2.drawContours(blank.copy(),[cn],-1,(255,255,255),-1)
        erode = cv2.erode(mask,kernel)

        # Find Corners of Block to crop image properly
        cropImg = cv2.bitwise_and(img,img,mask=erode)
        dst = cv2.cornerHarris(erode,25,11,0.03)
        ret, dst = cv2.threshold(dst,0.2*dst.max(),255,0)
        dst = np.uint8(dst)
        ret, labels, stats, centroids = cv2.connectedComponentsWithStats(dst)

        # Sort the corners to crop image via OpenCV standards
        df = pd.DataFrame(stats,columns=list(range(len(stats[0]))))
        # print(df)
        index = df[(df[4]<500)].index
        cent = np.array([centroids[i].round(0).astype(int) for i in index])
        sumCen = cent.sum(axis=1)
        x1,x2,y1,y2 = cent[np.argmin(sumCen)][0],cent[np.argmax(sumCen)][0],cent[np.argmin(sumCen)][1],cent[np.argmax(sumCen)][1]

        return cropImg[y1-3:y2+3,x1-3:x2+3]

    def findSticker(self,Proimg):
        hsv = cv2.cvtColor(Proimg.copy(),cv2.COLOR_BGR2HSV_FULL)
        hsv = cv2.bilateralFilter(hsv,50,15,15)
        # hsv = cv2.edgePreservingFilter(hsv,flags=1, sigma_s=60, sigma_r=0.4)
        # hsv = cv2.detailEnhance(hsv, sigma_s=10, sigma_r=0.15)
        mix = np.zeros(Proimg.shape[:2], np.uint8)

        for color,colRange in self.colDict.items():
            if "Tape" in color:
                thresh = cv2.inRange(hsv.copy(),colRange["LL"],colRange["UL"])
                thresh = cv2.morphologyEx(thresh,cv2.MORPH_CLOSE,np.ones((7,7),np.uint8))
                colProcessed = cv2.erode(thresh,np.ones((7,7),np.uint8))
                mix = cv2.bitwise_or(mix,colProcessed)
            else:
                thresh = cv2.inRange(hsv.copy(),colRange["LL"],colRange["UL"])
                thresh = cv2.erode(thresh,None)
                thresh = cv2.morphologyEx(thresh,cv2.MORPH_OPEN,None)
                cn,hi = cv2.findContours(thresh,cv2.RETR_TREE,cv2.CHAIN_APPROX_SIMPLE)
                for a,c in enumerate(cn):
                    colProcessed = np.zeros(thresh.shape[:2], np.uint8)
                    # if hi[0][a][2] == -1 and hi[0][a][3] == -1:
                    #     # continue
                    #     cv2.drawContours(colProcessed,[c],-1,color=(255,255,255),thickness=-1)
                    #     colProcessed = cv2.dilate(colProcessed,None)
                    # else:
                    cv2.drawContours(colProcessed,[c],-1,color=(255,255,255),thickness=-1)
                    colProcessed = cv2.dilate(colProcessed,None)
                    mix = cv2.bitwise_or(mix,colProcessed)

        return cv2.morphologyEx(mix,cv2.MORPH_CLOSE,np.ones((3,3),np.uint8),iterations=3)

    def colorSort(self,hsv,cnt):

        # Colour checking
        blank = np.zeros(hsv.shape[:2], np.uint8)
        cv2.drawContours(blank,[cnt],0,(255,255,255),-1)
        Stick_Mask = cv2.bitwise_and(hsv,hsv,mask=blank)

        Defects, hasCol = {}, {}
        for defects in self.defCode.keys(): Defects[defects] = 0
        for col in self.colDict:
            mask = cv2.inRange(Stick_Mask,self.colDict[col]['LL'],self.colDict[col]['UL'])
            hasCol[col] = np.sum(mask)
            if "Tape" in col and hasCol[col] < 30000: hasCol[col] = 0

        maxKey = max(hasCol,key=hasCol.get)
        maxVal = hasCol[maxKey]
        hasCol = hasCol.fromkeys(hasCol,0)
        hasCol[maxKey] = maxVal

        for colo,defe in self.defSticker.items():
            for i in hasCol:
                if colo == i: Defects[defe] += hasCol[i]
                else: Defects[defe] += 0

        return Defects, hasCol
    
    def TextAdd(self,Fimg,length,cnt,pieces,Textadd):
        if Textadd:
            approx = cv2.approxPolyDP(cnt, 0.0001*cv2.arcLength(cnt, True), True)
            n = approx.ravel()                                                               # Used to flatted the array containing the co-ordinates of the vertices.

            # Numbered Found stickers
            x = n[-2] + 5
            y = n[-1] - 5

            if Fimg.shape[0] < n[-2]*1.05: x = int(n[-2]*0.95)
            elif Fimg.shape[1] < n[-1]*1.05: y = int(n[-1]*0.95)
            elif n[-2] <= Fimg.shape[0]*0.05: x = int(Fimg.shape[0]*0.05)
            elif n[-1] <= Fimg.shape[1]*0.05: y = int(Fimg.shape[1]*0.05)

            cv2.putText(Fimg,f"{length:.0f}", (x, y),cv2.FONT_HERSHEY_COMPLEX, 0.5, (0, 255, 255), 1)
            # cv2.putText(Fimg,f"{pieces:.0f}", (n[-4]+5, n[-3]+30),cv2.FONT_HERSHEY_COMPLEX, 0.5, (0, 255, 255), 1)
            cv2.drawContours(Fimg,[approx],0,(255,255,255),1)
        return Fimg
    
    def accScan(self,Fimg,oddCol,Defects,cArea,cnt,value):
        for colName in oddCol:
            if oddCol[colName] > 0 and colName+"Area" in Defects.keys():
                Defects[colName+"Num"] += 1
                Defects[colName+"Area"] += value
            Fimg = self.TextAdd(Fimg,cArea,cnt,value,self.config['Trouble'])

        for colName in oddCol: Defects[colName+"Area"] = math.floor(Defects[colName+"Area"])

        return Fimg,Defects
    
    def chipPixelCnt(self,Proimg,cnts):

        Fimg = Proimg.copy()
        Defects = {}
        for colo in self.colDict:
            Defects[colo+"Num"] = 0
            Defects[colo+"Area"] = 0
        blur = cv2.GaussianBlur(Proimg.copy(), (5, 5), 0)
        hsv = cv2.cvtColor(blur,cv2.COLOR_BGR2HSV_FULL)

        for cnt in cnts:
            cArea = cv2.contourArea(cnt)
            if 1000000<cArea or cArea<30: continue

            realsize = cArea*self.avgPixLen
            pieces = math.floor(realsize/self.chipArea)
            length = math.sqrt(realsize)

            preDef,oddCol = self.colorSort(hsv,cnt)

            Fimg,Defects = self.accScan(Fimg,oddCol,Defects,cArea,cnt,realsize)

        print(Defects)
        # print(round(cArea,0))
        start = time.time()
        while True:
            cv2.namedWindow("Fimg",cv2.WINDOW_FREERATIO)
            cv2.imshow("Fimg",Fimg)
            k = cv2.waitKey(1) & 0xFF
            if time.time() - start > 3:
                break

        # print(f'Calibration Pin: {self.avgPixLen:.4f} mm^2/pixel')

        
# path = r"C:\Users\MES21106\Desktop\Block Auto Count\Code\Main\acc\May23\DMA"
path = r"C:\Users\MES21106\Desktop\Block Auto Count\Code\Main\block\May23\Eval 3-2\2350049100"

for j,i in enumerate(os.listdir(path)):
    print(f"Block {j+1}")
    start = time.time()
    Evaluation(cv2.imread(os.path.join(path,i)),chip='GJM03')
    print(time.time()-start)
    print("")