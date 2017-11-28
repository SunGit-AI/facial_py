import numpy as np
import cv

import Gui_OSDirectory

def f_test1():
    str_InputDir1 = ""
    face_cascade = cv2.CascadeClassifier('haarcascade_frontalface_default.xml')
    eye_cascade = cv2.CascadeClassifier('haarcascade_eye.xml')

    int_faceNotFound_counter = 0
    int_faceFound_counter = 0
    int_eyesFound_counter = 0
    listMatchedFiles = Gui_OSDirectory.OSDirectoryUtils.get_Dir_Files_with_Extention(str_InputDir1.replace('\\', '/'), '.xlsx')
    for str_File in listMatchedFiles:
        img = cv2.imread(str_File)
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        faces = face_cascade.detectMultiScale(gray, 1.3, 5)
        if len(faces)==0:
            int_faceNotFound_counter = int_faceNotFound_counter + 1
            print("witness: event: face not found")
        for (x,y,w,h) in faces:
            int_faceFound_counter = int_faceFound_counter + 1
            cv2.rectangle(img,(x,y),(x+w,y+h),(255,0,0),2)
            roi_gray = gray[y:y+h, x:x+w]
            roi_color = img[y:y+h, x:x+w]
            eyes = eye_cascade.detectMultiScale(roi_gray)
            for (ex,ey,ew,eh) in eyes: 
                int_eyesFound_counter = int_eyesFound_counter + 1
                cv2.rectangle(roi_color,(ex,ey),(ex+ew,ey+eh),(0,255,0),2)
            cv2.imshow('img',img)
            cv2.waitKey(0) #waitKey(0) will display the window infinitely until any keypress (it is suitable for image display)
            cv2.destroyAllWindows()