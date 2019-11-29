import logging

from fuzzywuzzy import process
import os

from wox import Wox, WoxAPI
from src import Registry, officeFile
from src.setting import Setting


# TODO Lowercase
class Main(Wox):
    MRUFiles = {}
    logging.basicConfig(filename='example.log', level=logging.DEBUG)

    def buildReg(self):
        self.MRUFiles = Registry.fetchRegistry()
        return

    def query(self, query):
        self.buildReg()
        returnResults = []

        if query=="":
            return self.showSetting()

        results = process.extract(query, self.MRUFiles.keys(), limit=4)
        for result in results:
            formattedPath = self.MRUFiles[result[0]]
            returnResults.append({
                "Title": result[0],
                "SubTitle": formattedPath,
                "IcoPath": officeFile.iconmatcher(result[0]),
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

    def showSetting(self):
        returnResults = []
        returnResults.append({
            "Title": "Office Recent File",
            "SubTitle": "",
            "IcoPath": "res//Icon.png"
        })
        returnResults.append({
            "Title": "Pinned file only",
            "SubTitle": str(Setting.getInstance().getPinned()),
            "IcoPath": "res//Icon.png",
            "JsonRPCAction": {
                "method": "pinned",
                "parameters": [],
                "dontHideAfterAction": False
            }
        })
        return returnResults

    def pinned(self):
        Setting.getInstance().setPinned()
        WoxAPI.show_msg("Searching only pinned file is set to "+str(Setting.getInstance().getPinned()),sub_title="")
        return

# Necessary code
if __name__ == "__main__":
    Main()
