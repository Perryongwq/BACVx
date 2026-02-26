import cv2
import math
import numpy as np

settings = {
"Pin Size": "5",
"CamResWidth": "1920",
"CamResHeight": "1080"
}

Black = {
    "LL": "0,1,80",
    "UL": "255,255,100"
}

Pin = {
    "LL": "0,1,105",
    "UL": "255,50,130"
}

class CaliTest():
    def __init__(self):
        super().__init__()
        self._initialize()
        imgArr = []
        self.cap = cv2.VideoCapture(0,cv2.CAP_DSHOW)
        self.cap.set(cv2.CAP_PROP_FRAME_WIDTH,int(settings['CamResWidth']))
        self.cap.set(cv2.CAP_PROP_FRAME_HEIGHT,int(settings['CamResHeight']))
        for i in range(10):
            image = self.cap.read()[1]
            imgArr.append(image[:,int(int(settings['CamResWidth'])/6):int(int(settings['CamResWidth'])/6*5)])
        self._main(imgArr)

    def __call__(self):
        return self.img,self.avgPixLen

    def _initialize(self):
        self.Area = (int(settings["Pin Size"])/2)**2*math.pi

    def _main(self,img):
        self.img, pixArea = self.houghCircle(img)
        
        self.avgPixLen = (self.Area/pixArea)*0.9
        print(f'Calibration Pin: {self.avgPixLen:.4f} mm/pixel')

    def houghCircle(self,imgArr):
        
        pixArr = []

        for img in imgArr:
            img = img[800:950,50:250]
            gray = cv2.cvtColor(img,cv2.COLOR_BGR2GRAY)
            blank = np.zeros(img.shape[:2], np.uint8)

            circles = cv2.HoughCircles(gray,cv2.HOUGH_GRADIENT,2,20,param1=20,param2=60,minRadius=3,maxRadius=13)
            print(circles)
            try: 
                for c in circles[0]:
                    cv2.circle(blank,(int(c[0]),int(c[1])),int(c[2]),(255,255,255),-1)

                circleMask = cv2.bitwise_and(img,img,mask=blank)
                cnts, _ = cv2.findContours(blank, cv2.RETR_EXTERNAL,cv2.CHAIN_APPROX_SIMPLE)

                blur = cv2.GaussianBlur(circleMask, (5, 5), 0)
                hsvimg = cv2.cvtColor(blur, cv2.COLOR_BGR2HSV_FULL)
                Black_LL = np.array([int(x) for x in Black['LL'].split(",")], dtype=np.uint8)
                Black_UL = np.array([int(y) for y in Black['UL'].split(",")], dtype=np.uint8)
                black = cv2.inRange(hsvimg,Black_LL,Black_UL)  

                for c in cnts:
                    cntArea = cv2.contourArea(c)
                    print(cntArea, np.sum(black))
                    if cntArea > 300 and np.sum(black) > 40000:
                        pixArr.append(cntArea)
            except: pass

        print(pixArr)
        pixArea = max(set(pixArr), key = pixArr.count) 

        return imgArr[pixArr.index(pixArea)],pixArea

CaliTest()