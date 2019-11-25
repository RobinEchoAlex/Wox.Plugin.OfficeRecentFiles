#Extract pinned file from registry

from winreg import *

def fetchRegistry():
        MRUFiles = {}
        wordKey = None

        reg = ConnectRegistry(None, HKEY_CURRENT_USER)
        wordKey = OpenKeyEx(reg,r"Software\Microsoft\Office\16.0\Word\User MRU\ADAL_F594B3B642FB17C41BF1285F133AD15D55BA76E8CB5C5CECD079FFED29C99780\File MRU")
        excelKey = OpenKeyEx(reg,r"Software\Microsoft\Office\16.0\Excel\User MRU\ADAL_F594B3B642FB17C41BF1285F133AD15D55BA76E8CB5C5CECD079FFED29C99780\File MRU")
        pptKey = OpenKeyEx(reg,r"Software\Microsoft\Office\16.0\PowerPoint\User MRU\ADAL_F594B3B642FB17C41BF1285F133AD15D55BA76E8CB5C5CECD079FFED29C99780\File MRU")

        i = 0
        while True:
            try:
                value = EnumValue(wordKey,i)
                if True:#TODO
                    fileInfo = extractInfo(value)
                    MRUFiles[fileInfo[0]] = fileInfo[1]
                i = i+1
            except OSError:  # EOF
                break
        return MRUFiles

def extractInfo(value):
    posOfAsterisk = value[1].find("*")
    filepath  = value[1][posOfAsterisk+1:]
    posOfSlash = filepath.rfind('\\')
    fileName = filepath[posOfSlash+1:]
    return fileName, filepath

def isPinned(value):
    if "F00000001" in value[1]:
        return True
    return False
