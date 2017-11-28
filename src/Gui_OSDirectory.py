'''
Created on 26.06.2015

@author: qxl7143
'''
import os
import glob



class OSDirectoryUtils(object):
    '''
    classdocs
    '''


    def __init__(self, params):
        '''
        Constructor
        '''

    @staticmethod
    def getAllDirectoryFilesByXExtension(strRootDiectroyIn, strGlobXExtensionIn):
        '''
        get all the matched files of the given directory with the given extension
        
        @return: the list of mateched files, [] if not found
        @rtype: list of file-paths
        '''
        listOutputMatchedFiles = []
        for dirName, subdirList, fileList in os.walk(strRootDiectroyIn):
            tmpListMatchedFiles = glob.glob1(dirName, strGlobXExtensionIn)
            for tmpListItemFile in tmpListMatchedFiles:
                tmpListItemFile = tmpListItemFile + '/' + tmpListItemFile
            listOutputMatchedFiles.extend(tmpListMatchedFiles)
        return listOutputMatchedFiles

    @staticmethod
    def get_Dir_Files_with_Extention(str_Src, str_Extention_In):  
        listOutputMatchedFiles = []
        for root, dirs, files in os.walk(str_Src):
            for file in files:
                if file.endswith(str_Extention_In):
                    str_Filefullpath = os.path.join(root, file)
                    listOutputMatchedFiles.append(str_Filefullpath)
        return listOutputMatchedFiles
    
if __name__ == '__main__':
    
    #listMatchedFiles = OSDirectoryUtils.getAllDirectoryFilesByXExtension('F:\tmp\xcels\E_oris', '*.xlsx')
    listMatchedFiles = OSDirectoryUtils.get_Dir_Files_with_Extention(r'F:\tmp\xcels\E_oris\\'.replace('\\', '/'), '.xlsx')
    for fname in listMatchedFiles:
        print(fname)