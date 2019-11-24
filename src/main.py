from fuzzywuzzy import process
import os
from random import random

from wox import Wox, WoxAPI
from src import Registry

# TODO Lowercase
class Main():
    MRUFiles = {}

    def buildReg(self):
        self.MRUFiles = Registry.fetchRegistry()
        return

    def query(self, query):
        if not self.MRUFiles:
            self.buildReg()
        returnResults = []
        results = process.extract(query, self.MRUFiles.keys(), limit=10)
        for result in results:
             formattedPath = self.MRUFiles[result[0]]
             returnResults.append({
                 "Title": result[0],
                 "SubTitle": formattedPath,
                 "IcoPath": "Res/word.ico",
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

# Necessary code
if __name__ == "__main__":
    Main()
