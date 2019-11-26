#Extract pinned file from registry

from winreg import *

from src.setting import Setting


def fetchRegistry():
        MRUFiles = {}
        wordKey = None

        reg = ConnectRegistry(None, HKEY_CURRENT_USER)

        wordKey = OpenKeyEx(reg,r"Software\Microsoft\Office\16.0\Word\User MRU\\"+ Setting.user + "\File MRU")
        excelKey = OpenKeyEx(reg,r"Software\Microsoft\Office\16.0\Excel\User MRU\\"+ Setting.user + "\File MRU")
        pptKey = OpenKeyEx(reg,r"Software\Microsoft\Office\16.0\PowerPoint\User MRU\\"+ Setting.user + "\File MRU")

        Keys = [wordKey,excelKey,pptKey]

        global_file_count = 0
        for key in Keys:
            count = 0
            while True:
                try:
                    value = EnumValue(key,count)
                    if True: #TODO
                        fileInfo = extractInfo(value)
                        MRUFiles[fileInfo[0]] = fileInfo[1]
                    global_file_count = global_file_count+1
                    count = count+1
                except OSError:  # EOF
                    break
        return MRUFiles

def createUserSet(reg):
    w_uKey = OpenKeyEx(reg, r"Software\Microsoft\Office\16.0\Word\User MRU")
    e_uKey = OpenKeyEx(reg, r"Software\Microsoft\Office\16.0\Excel\User MRU")
    p_uKey = OpenKeyEx(reg, r"Software\Microsoft\Office\16.0\PowerPoint\User MRU")
    uKeys = [w_uKey,e_uKey,p_uKey]

    for uKey in uKeys:
        num = QueryInfoKey(uKey)[0]
        for i in range(num):
            Setting.userSet.add(EnumKey(uKey, i))

def extractInfo(value):
    posOfAsterisk = value[1].find("*")
    filepath = value[1][posOfAsterisk+1:]
    posOfSlash = filepath.rfind('\\')
    fileName = filepath[posOfSlash+1:]
    return fileName, filepath

def isPinned(value):
    if "F00000001" in value[1]:
        return True
    return False
