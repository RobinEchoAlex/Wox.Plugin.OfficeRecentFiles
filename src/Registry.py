# Extract pinned file from registry

from winreg import *
from src.setting import Setting
import logging

MRUFiles = {}

def fetchRegistry():
    # if not bool(MRUFiles): not working normally
    #     return MRUFiles

    reg = ConnectRegistry(None, HKEY_CURRENT_USER)
    createUserSet(reg)
    setting = Setting.getInstance()

    logging.basicConfig(filename='example.log', level=logging.DEBUG)

    for app in setting.userDict.keys():
        appName = ""
        if (app == "w_user"):
            appName = "Word"
        elif (app == "e_user"):
            appName = "Excel"
        elif (app == "p_user"):
            appName = "PowerPoint"
        userSet = setting.userDict.get(app)
        for user in userSet:
            key = OpenKeyEx(reg,
                            "Software\\Microsoft\\Office\\16.0\\" + appName + "\\User MRU\\" + user + "\\File MRU")
            num = QueryInfoKey(key)[1] #the number of values that this key has
            #logging.debug(num)
            for i in range(num):
                value = EnumValue(key,i)
                #logging.debug(value)
                logging.debug(setting.getPinned())
                if setting.getPinned() and isPinned(value) or not setting.getPinned():
                    fileInfo = extractInfo(value)
                    MRUFiles[fileInfo[0]] = fileInfo[1]
    #logging.debug(MRUFiles)
    return MRUFiles


def createUserSet(reg):
    w_uKey = OpenKeyEx(reg, r"Software\Microsoft\Office\16.0\Word\User MRU")
    e_uKey = OpenKeyEx(reg, r"Software\Microsoft\Office\16.0\Excel\User MRU")
    p_uKey = OpenKeyEx(reg, r"Software\Microsoft\Office\16.0\PowerPoint\User MRU")
    uKeys = [w_uKey, e_uKey, p_uKey]

    for uKey in uKeys:
        num = QueryInfoKey(uKey)[0]
        for i in range(num):
            if uKey == w_uKey:
                Setting.userDict.get("w_user").add(EnumKey(uKey, i))
            elif uKey == e_uKey:
                Setting.userDict.get("e_user").add(EnumKey(uKey, i))
            elif uKey == p_uKey:
                Setting.userDict.get("p_user").add(EnumKey(uKey, i))


def extractInfo(value):
    posOfAsterisk = value[1].find("*")
    filepath = value[1][posOfAsterisk + 1:]
    posOfSlash = filepath.rfind('\\')
    fileName = filepath[posOfSlash + 1:]
    return fileName, filepath

def isPinned(value):
    if "F00000001" in value[1]:
        return True
    return False
