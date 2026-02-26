import cv2
import numpy as np

# Parameters to change
####################################################################################################################################################

cb = {}
parameter = ['dp','minc','para1','para2','minr','maxr']

####################################################################################################################################################

# Trackbar callback function to update HSV values
def callback(x):
    for para in parameter:
        cb[para] = cv2.getTrackbarPos(para,'controls')

# Create seperate windows for trackbar, mask and image
winName = ['controls','res','mask']
for name in winName:
    cv2.namedWindow(name,cv2.WINDOW_FREERATIO)
    cv2.resizeWindow(name,500,500)

for p in parameter: 
    cv2.createTrackbar(p,'controls',1,200,callback)
    cb[p] = cv2.getTrackbarPos(p,'controls')

cap = cv2.VideoCapture(0,cv2.CAP_DSHOW)
cap.set(cv2.CAP_PROP_FRAME_WIDTH,1920)
cap.set(cv2.CAP_PROP_FRAME_HEIGHT,1080)

while True:

    _, img = cap.read()
    img = img[800:950,50:250]
    blank = np.zeros(img.shape[:2], np.uint8)

    pin = cv2.cvtColor(img,cv2.COLOR_BGR2GRAY)

    circles = cv2.HoughCircles(pin,cv2.HOUGH_GRADIENT,cb['dp'],cb['minc'],param1=cb['para1'],param2=cb['para2'],minRadius=cb['minr'],maxRadius=cb['maxr'])
    try:
        for c in circles[0]:
            cv2.circle(blank,(int(c[0]),int(c[1])),int(c[2]),(255,255,255),-1)
    except: blank

    img = cv2.bitwise_and(pin,pin,mask=blank)

    cv2.imshow("res",img)
    cv2.imshow("mask",blank)

    if cv2.waitKey(1) & 0xFF == ord('q'):
        print(cb)
        break
  
# After the loop release the cap object
cap.release()
# Destroy all the windows
cv2.destroyAllWindows()