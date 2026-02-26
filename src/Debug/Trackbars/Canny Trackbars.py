import cv2  
import numpy as np

# Parameters to change
####################################################################################################################################################

parameter = ['blur','low thresh','high thresh']
blurList = ['Blur','Median','Gaussian','Bilateral']
cb = {}

blurmethod = blurList[3]

####################################################################################################################################################

def blurType(img):
    if blurmethod == 'Blur': return cv2.blur(img, (bKer, bKer))
    elif blurmethod == 'Median': return cv2.medianBlur(img, bKer)
    elif blurmethod == 'Gaussian': return cv2.GaussianBlur(img, (bKer, bKer), 0)
    elif blurmethod == 'Bilateral': return cv2.bilateralFilter(img, 15, bKer, bKer)

# Trackbar callback function to update HSV values
def callback(x):
    for para in parameter:
        cb[para] = cv2.getTrackbarPos(para,'controls')

# Create seperate windows for trackbar, mask and image
winName = ['controls','res','blur']
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
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    bKer = cb['blur']+1 if cb['blur']%2 == 0 else cb['blur']
    blur = blurType(gray)

    canny = cv2.Canny(blur,cb['low thresh'],cb['high thresh'],3)
    dilate = cv2.dilate(canny,np.ones((5,5),np.uint8))

    #show image
    cv2.imshow('res',canny)
    cv2.imshow('blur',blur)

    #waitfor the user to press escape and break the while loop 
    k = cv2.waitKey(1) & 0xFF
    if k == 27:
        break
		
#destroys all window
cv2.destroyAllWindows()