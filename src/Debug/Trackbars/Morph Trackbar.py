import cv2  
import numpy as np

# Parameters to change
####################################################################################################################################################

parameter = ['LKernel','RKernel']
morphList = ['Erode','Dilate','Open','Close','Rectangle','Cross']
cb = {}

morphmethod = morphList[2]

####################################################################################################################################################

def morphType(img):
    kernel = np.ones((LKer, RKer),np.uint8)
    if morphmethod == 'Erode': return cv2.erode(img, kernel)
    elif morphmethod == 'Dilate': return cv2.dilate(img, kernel)
    elif morphmethod == 'Open': return cv2.morphologyEx(img, cv2.MORPH_OPEN, kernel)
    elif morphmethod == 'Close': return cv2.morphologyEx(img, cv2.MORPH_CLOSE, kernel)
    elif morphmethod == 'Rectangle': return cv2.morphologyEx(img, cv2.MORPH_RECT, kernel)
    elif morphmethod == 'Cross': return cv2.morphologyEx(img, cv2.MORPH_CROSS, kernel)

# Trackbar callback function to update HSV values
def callback(x):
    for para in parameter:
        cb[para] = cv2.getTrackbarPos(para,'controls')

# Create seperate windows for trackbar, mask and image
winName = ['controls','mask','res']
for name in winName:
    cv2.namedWindow(name,cv2.WINDOW_FREERATIO)
    cv2.resizeWindow(name,500,500)

# Create trackbars for Canny
for p in parameter: 
    cv2.createTrackbar(p,'controls',0,255,callback)
    cb[p] = cv2.getTrackbarPos(p,'controls')

cap = cv2.VideoCapture(0,cv2.CAP_DSHOW)
cap.set(cv2.CAP_PROP_FRAME_WIDTH,1920)
cap.set(cv2.CAP_PROP_FRAME_HEIGHT,1080)

while True:

    _, img = cap.read()
    img = img[:,int(1920/6):int(1920/6*5)]

    blank = np.zeros(img.shape[:2], np.uint8)
    gray = cv2.cvtColor(img,cv2.COLOR_BGR2GRAY)
    blur = cv2.bilateralFilter(cv2.cvtColor(img.copy(),cv2.COLOR_BGR2GRAY),25,10,10)
    thresh = cv2.adaptiveThreshold(blur,255,cv2.ADAPTIVE_THRESH_GAUSSIAN_C,cv2.THRESH_BINARY_INV,17,3)
    dilate = cv2.dilate(thresh,np.ones((5,5),np.uint8))
    cnt, hier = cv2.findContours(dilate,cv2.RETR_EXTERNAL,cv2.CHAIN_APPROX_SIMPLE)
    cn = sorted(cnt, key=lambda x: cv2.contourArea(x))[-1]
    thresh = cv2.drawContours(blank.copy(),[cn],-1,(255,255,255),-1)
    thresh = cv2.erode(thresh,np.ones((5,5),np.uint8))

    LKer = cb['LKernel']+1 if cb['LKernel']%2 == 0 else cb['LKernel']
    RKer = cb['RKernel']+1 if cb['RKernel']%2 == 0 else cb['RKernel']
    morph = morphType(thresh)

    img = cv2.bitwise_and(img,img,mask=morph)

    #show image
    cv2.imshow('mask',morph)
    cv2.imshow('res',img)

    #waitfor the user to press escape and break the while loop 
    k = cv2.waitKey(1) & 0xFF
    if k == 27:
        break
		
#destroys all window
cv2.destroyAllWindows()