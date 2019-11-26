import configparser

class Setting:
    userSet = set()
    __user = ""
    __enable = {"word":True,"excel":True,"ppt":True}
    __pinned = False


    def setUser(self,user):
        if user not in Setting.userSet:
            raise Exception("Requested user is not found")
        Setting.__user = user
        Setting.writeSetting(self)

    def getUser(self):
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
        __user = config['user']
        __enable = config['enable']
        __pinned = config.getboolean('pinned')
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