import cv2  
import numpy as np

# Parameters to change
####################################################################################################################################################

path = r"C:\Users\MES21106\Desktop\Block Auto Count\Code\Main\block\May23\2340260200\2340260200_02-05-23_154942.png"
parameter = ['flag','sigma_s','sigma_r']
filter = ['Edge','Detail']
cb = {}

filterMethod = filter[0]

####################################################################################################################################################

def filterType(img):
    if filterMethod == 'Edge': return cv2.edgePreservingFilter(img, flags=flag, sigma_r=sigma_r, sigma_s=sigma_s)
    elif filterMethod == 'Detail': return cv2.detailEnhance(img, sigma_r=sigma_r, sigma_s=sigma_s)

# Trackbar callback function to update HSV values
def callback(x):
    for para in parameter:
        cb[para] = cv2.getTrackbarPos(para,'controls')

# Create seperate windows for trackbar, res and image
winName = ['controls','img','res']
for name in winName:
    cv2.namedWindow(name,cv2.WINDOW_FREERATIO)
    cv2.resizeWindow(name,500,500)

# Create trackbars for Canny
for p in parameter: 
    cv2.createTrackbar(p,'controls',0,255,callback)
    cb[p] = cv2.getTrackbarPos(p,'controls')

while True:

    img = cv2.imread(path)
    flag = cb['flag']
    sigma_r = cb['sigma_r']
    sigma_s = cb['sigma_s']/100

    res = filterType(img)
    mask = cv2.inRange(cv2.cvtColor(res,cv2.COLOR_BGR2HSV_FULL),(20,45,200),(35,100,255))
    img = cv2.bitwise_and(img,img,mask=mask)

    #show image
    cv2.imshow('img',img)
    cv2.imshow('res',res)

    #waitfor the user to press escape and break the while loop 
    k = cv2.waitKey(1) & 0xFF
    if k == 27:
        break
		
#destroys all window
cv2.destroyAllWindows()