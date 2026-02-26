import cv2  
import numpy as np

# Parameters to change
####################################################################################################################################################

low = ['Low R','Low G','Low B']
high = ['High R','High G','High B']
blurList = ['Blur','Median','Gaussian','Bilateral']
cb = {}

blurmethod = blurList[3]

toBlur = True

####################################################################################################################################################

# Return blur based on selection
def blurType(img):
    if blurmethod == 'Blur': return cv2.blur(img, (bKer, bKer))
    elif blurmethod ==  'Median': return cv2.medianBlur(img, bKer)
    elif blurmethod ==  'Gaussian': return cv2.GaussianBlur(img, (bKer, bKer), 3)
    elif blurmethod ==  'Bilateral': return cv2.bilateralFilter(img, 15, bKer, bKer)

# Trackbar callback function to update HSV values
def callback(x):
    for i in range(len(low)):
        cb[low[i]] = cv2.getTrackbarPos(low[i],'controls')
        cb[high[i]] = cv2.getTrackbarPos(high[i],'controls')
    cb['blur'] = cv2.getTrackbarPos('blur','controls')

# Create seperate windows for trackbar, mask and image
winName = ['controls','mask','res']
for name in winName:
    cv2.namedWindow(name,cv2.WINDOW_FREERATIO)
    cv2.resizeWindow(name,500,500)

# Create trackbars for high & low [R,G,B]
for i in range(len(low)):
    cv2.createTrackbar(low[i],'controls',0,255,callback)
    cv2.createTrackbar(high[i],'controls',255,255,callback)
    cb[low[i]] = cv2.getTrackbarPos(low[i],'controls')
    cb[high[i]] = cv2.getTrackbarPos(high[i],'controls')
cv2.createTrackbar('blur','controls',0,255,callback)
cb['blur'] = cv2.getTrackbarPos('blur','controls')

cap = cv2.VideoCapture(0,cv2.CAP_DSHOW)
cap.set(cv2.CAP_PROP_FRAME_WIDTH,1920)
cap.set(cv2.CAP_PROP_FRAME_HEIGHT,1080)

while True:

    _, img = cap.read()
    img = img[:,int(1920/6):int(1920/6*5)]
    rgb = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
    bKer = cb['blur']+1 if cb['blur']%2 == 0 else cb['blur']
    blur = blurType(rgb) if toBlur else rgb

    larr,harr = [],[]
    for l in low: larr.append(cb[l])
    for h in high:harr.append(cb[h])

    rgb_low = np.array(larr, np.uint8)
    rgb_high = np.array(harr, np.uint8)

    mask = cv2.inRange(rgb, rgb_low, rgb_high)
    res = cv2.bitwise_and(img, img, mask=mask)

    cv2.imshow('mask',mask)
    cv2.imshow('res',res)

    # Break loop by pressing escape button
    k = cv2.waitKey(1) & 0xFF
    if k == 27:
        break
		
cv2.destroyAllWindows()