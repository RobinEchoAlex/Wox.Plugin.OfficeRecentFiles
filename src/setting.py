import configparser

#TODO static
class Setting:
    __instance = None

    userSet = set()
    __user = ""
    __enable = {"word":True,"excel":True,"ppt":True}
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
        return Setting.__instance

    def setUser(self,user):
        if user not in Setting.userSet:
            raise Exception("Requested user is not found")
        Setting.__user = user
        Setting.writeSetting(self)

    def getUser(self):
        if (Setting.__user ==""):
            Setting.__user = Setting.userSet.pop()
            Setting.userSet.add(Setting.__user)
        return Setting.__user

    def setEnable(self,app):
        if app not in Setting.__enable.keys():
            raise Exception("Requested app is not existed")
        Setting.__enable[app] = not Setting.__enable[app]
        Setting.writeSetting(self)

    def getEnable(self,app):
        if app not in Setting.__enable.keys():
            raise Exception("Requested app is not existed")
        return Setting.__enable[app]

    def setPinned(self,app):
        Setting.__pinned = not Setting.__pinned
        return

    def getPinned(self):
        return Setting.__pinned

    def loadSetting(self):
        config = configparser.ConfigParser()
        config .read('config.ini')

        app = config['APP']
        __user = app['user']
        __pinned = app.getboolean('pinned')

        Setting.__enable['word']=config.getboolean('ENABLE','word')
        Setting.__enable['excel'] = config.getboolean('ENABLE','excel')
        Setting.__enable['ppt'] = config.getboolean('ENABLE','ppt')
        return

    def writeSetting(self):
        config = configparser.ConfigParser()
        config['APP'] = {}
        app = config['APP']
        app['user'] = Setting.__user
        app['pinned'] = str(Setting.__pinned)

        config['ENABLE']={}
        ena = config['ENABLE']
        ena['word'] = str(Setting.__enable.get('word'))
        ena['excel'] = str(Setting.__enable.get('excel'))
        ena['ppt'] = str(Setting.__enable.get('ppt'))

        with open('config.ini','w') as configFile:
            config.write(configFile)
        return