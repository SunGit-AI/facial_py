'''
Created on 22.11.2017

@author: qi11028
'''
import cv2

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

if __name__ == '__main__':
    oVideo_Utils = Video_Utils()
    oVideo_Utils.f_ViedeoToImages(Video_Utils.clsStr1.replace('\\', '/'))
    pass