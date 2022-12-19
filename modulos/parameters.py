import json
from modulos.crypto import Crypto
from os import getcwd

class Parameters:
    def __init__(self, key):
        self.c = Crypto(key)
        self.__sapEnv = None
        self.__userId = None
        self.__userPw = None
# -------------------------------------------------------
    @property
    def sapEnv(self):
        return self.__sapEnv

    @sapEnv.setter
    def sapEnv(self, value):
        self.__sapEnv = value
# -------------------------------------------------------
    @property
    def userId(self):
        return self.c.decrypt(self.__userId).decode('ascii')

    @userId.setter
    def userId(self, value):
        self.__userId = self.c.crypt(value)
# -------------------------------------------------------
    @property
    def userPw(self):
        return self.c.decrypt(self.__userPw).decode('ascii')

    @userPw.setter
    def userPw(self, value):
        self.__userPw = self.c.crypt(value)
# -------------------------------------------------------
    def writeParametersFile(self):
        par = {
            'sapEnv': self.__sapEnv,
            'userId': self.__userId.decode('ascii'),
            'userPw': self.__userPw.decode('ascii'),
        }

        par = json.dumps(par)

        f = open(getcwd() + '\parameters.json', 'w', encoding='utf-8')
        f.write(par)
        f.close
# -------------------------------------------------------
    def readParameters(self, file=getcwd() + '\parameters.json'):
        f = open(file, 'r', encoding='utf-8')
        par = json.load(f)
        f.close()

        self.__sapEnv = par['sapEnv']
        self.__userId = par['userId'].encode('ascii')
        self.__userPw = par['userPw'].encode('ascii')