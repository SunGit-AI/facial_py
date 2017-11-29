'''
Created on 22.11.2017

@author: qi11028
'''
import cv2
import numpy as np
from Gui_OSDirectory import *

class Video_Utils(object):
    clsStr1 = r"F:\SunsNoteBook\Data\20171122_141342.mp4"
    clsStr2 = r"F:\SunsNoteBook\Data\out\\"
    def __init__(self):
        pass
    
    def f_ViedeoToImages(self, str_File_In):
        vidcap = cv2.VideoCapture(str_File_In)
        success,image = vidcap.read()
        count = 0
        success = True
        while success:
          success,image = vidcap.read()
          print('Read a new frame: ', success)
          if image != None:
              tuple_shape = image.shape
              rows = tuple_shape[0]
              cols = tuple_shape[1]
              oRotateMatrix = cv2.getRotationMatrix2D((cols/2,rows/2),90,1) # rotates the image by 90 degree with respect to center without any scaling. 
              orotate_image = cv2.warpAffine(image,oRotateMatrix,(cols,rows))
              cv2.imwrite(Video_Utils.clsStr2 + "frame%d.jpg" % count, orotate_image)     # save frame as JPEG file
              count += 1    
              
class Facial_Reg(object):
    clsStr_ClassifierFace = r"F:\SunsNoteBook\Data\classifiers\20046\haarcascades\haarcascade_frontalface_default.xml"
    clsStr_ClassifierEye = r"F:\SunsNoteBook\Data\classifiers\20046\haarcascades\haarcascade_eye.xml"
    clsStr_ClassifierEye_1 = r"F:\SunsNoteBook\Data\classifiers\20046\haarcascades\haarcascade_eye_tree_eyeglasses.xml"
    clsStr_ClassifierEye_2 = r"F:\SunsNoteBook\Data\classifiers\20046\haarcascades_cuda\haarcascade_eye_tree_eyeglasses.xml"
    clsStr_ClassifierEye_3 = r"F:\SunsNoteBook\Data\classifiers\cascades\haarcascade_mcs_eyepair_small.xml"
    
    def __init__(self):
        pass
    
    def test_f_1(self):
        face_cascade = cv2.CascadeClassifier(Facial_Reg.clsStr_ClassifierFace.replace('\\', '/'))
        eye_cascade = cv2.CascadeClassifier(Facial_Reg.clsStr_ClassifierEye_3.replace('\\', '/'))

        listMatchedFiles = OSDirectoryUtils.get_Dir_Files_with_Extention(Video_Utils.clsStr2.replace('\\', '/'), '.jpg')
        for fname in listMatchedFiles:
            print(fname)
            img = cv2.imread(fname)
            gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
            faces = face_cascade.detectMultiScale(gray) #1.3, 5
            for (x,y,w,h) in faces:
                cv2.rectangle(img,(x,y),(x+w,y+h),(255,0,0),2)
                roi_gray = gray[y:y+h, x:x+w]
                roi_color = img[y:y+h, x:x+w]
                eyes = eye_cascade.detectMultiScale(roi_gray)
                for (ex,ey,ew,eh) in eyes:
                    cv2.rectangle(roi_color,(ex,ey),(ex+ew,ey+eh),(0,255,0),2)
            cv2.imshow('img',img)
            cv2.waitKey(0)
            cv2.destroyAllWindows() #
            
    def test_f_2(self):
        face_cascade = cv2.CascadeClassifier('haarcascade_frontalface_default.xml') 
        eye_cascade = cv2.CascadeClassifier('haarcascade_eye.xml')
        
        img = cv2.imread('sachin.jpg') 
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY) #@endcode Now we find the faces in the image. If faces are found, it returns the positions of detected faces as Rect(x,y,w,h). Once we get these locations, we can create a ROI for the face and apply eye detection on this ROI (since eyes are always on the face !!! ). @code{.py} 
        faces = face_cascade.detectMultiScale(gray, 1.3, 5) 
        for (x,y,w,h) in faces: 
            cv2.rectangle(img,(x,y),(x+w,y+h),(255,0,0),2) 
            roi_gray = gray[y:y+h, x:x+w] 
            roi_color = img[y:y+h, x:x+w] 
            eyes = eye_cascade.detectMultiScale(roi_gray) 
            for (ex,ey,ew,eh) in eyes: 
                cv2.rectangle(roi_color,(ex,ey),(ex+ew,ey+eh),(0,255,0),2)
        
        cv2.imshow('img',img) 
        cv2.waitKey(0) 
        cv2.destroyAllWindows() 

if __name__ == '__main__':
    #oVideo_Utils = Video_Utils()
    #oVideo_Utils.f_ViedeoToImages(Video_Utils.clsStr1.replace('\\', '/'))
    oFacial_Reg = Facial_Reg()
    oFacial_Reg.test_f_1()
    
    pass