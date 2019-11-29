import configparser
import logging

# TODO static
class Setting:
    __instance = None
    logging.basicConfig(filename='example.log', level=logging.DEBUG)

    userDict = {"w_user":set(), "e_user":set(), "p_user":set()}
    __pinned = False

    def __init__(self):
        if Setting.__instance != None:
            raise Exception("User getInstance() instead")
        else:
            Setting.__instance = self

    @staticmethod
    def getInstance():
        if Setting.__instance == None:
            Setting()
        Setting.__instance.loadSetting()
        return Setting.__instance

    def setPinned(self):
        self.__pinned = not self.__pinned
        logging.debug("Pinned in set "+str(self.__pinned))
        self.writeSetting()
        return

    def getPinned(self):
        logging.debug("Pinned in get "+str(self.__pinned))
        return self.__pinned

    def loadSetting(self):
        config = configparser.ConfigParser()
        config.read('config.ini')

        app = config['APP']
        self.__pinned = app.getboolean('pinned')

        logging.debug("Pinned in read "+str(self.__pinned))
        return

    def writeSetting(self):
        config = configparser.ConfigParser()
        config['APP'] = {}
        app = config['APP']
        app['pinned'] = str(self.__pinned)

        logging.debug("Pinned in write "+str(self.__pinned))

        with open(r'config.ini', 'w') as configFile:
            config.write(configFile)
        return
