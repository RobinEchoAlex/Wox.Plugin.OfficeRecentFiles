# dev might want different icon for different type of file in the future
extension_icon = {
    "doc": "res//word.ico",
    "dot": "res//word.ico",
    "wbk": "res//word.ico",
    "docx": "res//word.ico",
    "docm": "res//word.ico",
    "dotx": "res//word.ico",
    "dotm": "res//word.ico",
    "docb": "res//word.ico",

    "xls": "res//excel.ico",
    "xlt": "res//excel.ico",
    "xlm": "res//excel.ico",
    "xlsx": "res//excel.ico",
    "xlsm": "res//excel.ico",
    "xltx": "res//excel.ico",
    "xltm": "res//excel.ico",
    "xlsb": "res//excel.ico",
    "xla": "res//excel.ico",
    "xlam": "res//excel.ico",
    "xll": "res//excel.ico",
    "xlw": "res//excel.ico",
    #TODO csv

    "ppt": "res//ppt.ico",
    "pot": "res//ppt.ico",
    "pps": "res//ppt.ico",
    "pptx": "res//ppt.ico",
    "pptm": "res//ppt.ico",
    "potx": "res//ppt.ico",
    "potm": "res//ppt.ico",
    "sldx": "res//ppt.ico"
}


def iconmatcher(fileName):
    pos = fileName.rfind(".")
    extension = fileName[pos + 1:]
    if extension_icon.__contains__(extension):
        icon = extension_icon[extension]
    else:
        icon = "res//Icon.png"
    return icon
