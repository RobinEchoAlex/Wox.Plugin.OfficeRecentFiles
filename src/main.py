from fuzzywuzzy import process
import os

from wox import Wox, WoxAPI
from src import Registry, file
from src.setting import Setting


# TODO Lowercase
class Main(Wox):
    MRUFiles = {}

    def buildReg(self):
        self.MRUFiles = Registry.fetchRegistry()
        return

    def query(self, query):


        self.buildReg()
        returnResults = []

        if query=="":
            return self.setting()

        results = process.extract(query, self.MRUFiles.keys(), limit=4)
        for result in results:
            formattedPath = self.MRUFiles[result[0]]
            returnResults.append({
                "Title": result[0],
                "SubTitle": formattedPath,
                "IcoPath": file.iconmatcher(result[0]),
                "JsonRPCAction": {
                    "method": "openFile",
                    "parameters": [formattedPath],
                    "dontHideAfterAction": False
                }
            })
        return returnResults

    def openFile(self, filePath):
        os.startfile(filePath)
        # TODO how about a file is not longer existed

    def setting(self):
        returnResults = []
        returnResults.append({
            "Title": "Office Recent File",
            "SubTitle": "",
            "IcoPath": "res//Icon.png"
        })
        returnResults.append({
            "Title": "User Setting",
            "SubTitle": "configure the user whose most recent files to show",
            "IcoPath": "res//Icon.png",
            "JsonRPCAction": {
            "method": "userSetting",
            "parameters": [],
            "dontHideAfterAction": False
        }
        })

    def userSetting(self):
        setting = Setting.getInstance()
        users = setting.getUser()


# Necessary code
if __name__ == "__main__":
    Main()
