import cv2  
import numpy as np

# Parameters to change
####################################################################################################################################################

cb = {}
parameter = ['blur','thresh']
thresList = ['Binary','Binary_Inv','Trunc','ToZero','ToZero_Inv','Otsu','BinOtsu']
blurList = ['Blur','Median','Gaussian','Bilateral']
techList = ['Threshold','Adaptive']
adaptList = ['Mean','Gaussian']

thresmethod = thresList[1]
blurmethod = blurList[3]
Techmethod = techList[1]
adaptmethod = adaptList[1]

equalize = False
contrast = [False,2.5]      # clipLimit
toBlur = True

####################################################################################################################################################

# Equalize Histogram
def hist(img):

    # Histogram takes a list and calculate the number of occurence based on the bins and range.
    # Bins is the number of bins you would like to divide the range into
    # If no range is provided, the range is between the min and max of list
    # For images, pixel value is from 0 to 255, thus 256 values
    hist,bins = np.histogram(img.copy().flatten(),256,[0,256])
    cumsum = hist.cumsum()                                                          # Cumulative Sum, always positive
    nonzero = np.ma.masked_equal(cumsum,0)                                          # Remove 0 elements in list
    equalize = (nonzero - nonzero.min())/(nonzero.max()-nonzero.min())*255          # Equalize elements in list that are higher in occurence
    df = np.ma.filled(equalize,0).astype('uint8')                                   # Fill blank with 0 elements in list
    eHist = df[img]                                                                 # Replaces image's pixel with equalized elements

    return eHist

# Improve Contrast
def CLAHE(img,clip):
    lab = cv2.cvtColor(img, cv2.COLOR_BGR2Lab)
    l_channel, a, b = cv2.split(lab)
    clahe = cv2.createCLAHE(clipLimit=clip)
    image = clahe.apply(l_channel)

    return image

# Return blur based on selection
def blurType(img):
    if blurmethod == 'Blur': return cv2.blur(img, (bKer, bKer))
    elif blurmethod ==  'Median': return cv2.medianBlur(img, bKer)
    elif blurmethod ==  'Gaussian': return cv2.GaussianBlur(img, (bKer, bKer), 11)
    elif blurmethod ==  'Bilateral': return cv2.bilateralFilter(img, 15, bKer, bKer)

# Return Threshold types based on selection
def thresType(img):
    if thresmethod == 'Binary': thres = cv2.THRESH_BINARY
    elif thresmethod == 'Binary_Inv': thres = cv2.THRESH_BINARY_INV
    elif thresmethod == 'Trunc': thres = cv2.THRESH_TRUNC
    elif thresmethod == 'ToZero': thres = cv2.THRESH_TOZERO
    elif thresmethod == 'ToZero_Inv': thres = cv2.THRESH_TOZERO_INV
    elif thresmethod == 'Otsu': thres = cv2.THRESH_OTSU
    elif thresmethod == 'BinOtsu': thres = cv2.THRESH_BINARY+cv2.THRESH_OTSU

    if Techmethod == 'Threshold': return cv2.threshold(img,cb['thresh'],255,thres)[1]
    elif Techmethod == 'Adaptive': 
        tVal = cb['thresh']+1 if cb['thresh']%2 == 0 else cb['thresh']
        tVal = 3 if cb['thresh'] < 3 else tVal
        if adaptmethod == 'Mean': return cv2.adaptiveThreshold(img,255,cv2.ADAPTIVE_THRESH_MEAN_C,thres,tVal,3)
        elif adaptmethod == 'Gaussian': return cv2.adaptiveThreshold(img,255,cv2.ADAPTIVE_THRESH_GAUSSIAN_C,thres,tVal,3)

# Trackbar callback function to update values
def callback(x):
    for para in parameter:
        cb[para] = cv2.getTrackbarPos(para,'controls')

# Create seperate windows for trackbar, mask and image
winName = ['controls','res','mask']
for name in winName:
    cv2.namedWindow(name,cv2.WINDOW_FREERATIO)
    cv2.resizeWindow(name,500,500)

# Create trackbars for Threshold
for p in parameter: 
    cv2.createTrackbar(p,'controls',0,255,callback)
    cb[p] = cv2.getTrackbarPos(p,'controls')

cap = cv2.VideoCapture(0,cv2.CAP_DSHOW)
cap.set(cv2.CAP_PROP_FRAME_WIDTH,1920)
cap.set(cv2.CAP_PROP_FRAME_HEIGHT,1080)

while True:

    _, img = cap.read()
    img = img[:,int(1920/6):int(1920/6*5)]
    if contrast[0]: gray = CLAHE(img,contrast[1])
    if equalize: gray = hist(img)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    bKer = cb['blur']+1 if cb['blur']%2 == 0 else cb['blur']
    blur = blurType(gray) if toBlur else gray
    mask = thresType(blur)

    res = cv2.bitwise_and(img,img,mask=mask)

    #show image
    cv2.imshow('mask',mask)
    cv2.imshow('res',res)

    #waitfor the user to press escape and break the while loop 
    k = cv2.waitKey(1) & 0xFF
    if k == 27:
        break
		
#destroys all window
cv2.destroyAllWindows()