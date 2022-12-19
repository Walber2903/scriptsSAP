from os import getcwd
# Create a dictionary with parameters and generate a crypto key
# Example:
# from modulos.createParameters import CreateParameters
# dictPar = dict()
# dictPar['sapEnv'] = ('nc', 'SAP Production')
# dictPar['userId'] = ('c', 'myUserID')
# dictPar['userPw'] = ('c', 'myUserPW')
# dictPar['path'] = ('nc', r'C:\Users\MarcioRibeiro\Documents\Reports\\')
# c = CreateParameters(dictPar, 'myCryptoKey')

class CreateParameters:
    def __init__(self, dictParameters, key) -> None:
        self.dictParameters = dictParameters
        self.key = key
        self.mountModuleParameters()
        self.writeInFile()
        self.mountModuleWriteParameters()
        self.writeInFile()

    def mountModuleParameters(self):
        self.file = getcwd() + '\\modulos\\parameters.py'

        self.text = 'import json' + chr(10)
        self.text += 'from modulos.crypto import Crypto' + chr(10)
        self.text += 'from os import getcwd' + chr(10) + chr(10)
        self.text += 'class Parameters:' + chr(10)
        self.text += '    def __init__(self, key):' + chr(10)
        self.text += '        self.c = Crypto(key)' + chr(10)

        for key in self.dictParameters:
            self.text += f'        self.__{key} = None' + chr(10)

        for key, item in self.dictParameters.items():
            self.text += '# -------------------------------------------------------' + chr(10)
            self.text += '    @property' + chr(10)
            self.text += f'    def {key}(self):' + chr(10)
            if item[0] == 'c':
                self.text += f"        return self.c.decrypt(self.__{key}).decode('ascii')" + chr(10) + chr(10)
                self.text += f'    @{key}.setter' + chr(10)
                self.text += f'    def {key}(self, value):' + chr(10)
                self.text += f'        self.__{key} = self.c.crypt(value)' + chr(10)
            else:
                self.text += f'        return self.__{key}' + chr(10) + chr(10)
                self.text += f'    @{key}.setter' + chr(10)
                self.text += f'    def {key}(self, value):' + chr(10)
                self.text += f'        self.__{key} = value' + chr(10)

        self.text += '# -------------------------------------------------------' + chr(10)
        self.text += '    def writeParametersFile(self):' + chr(10)
        self.text += '        par = {'

        for key, item in self.dictParameters.items():
            if item[0] == 'c':
                self.text += chr(10) + f"            '{key}': self.__{key}.decode('ascii'),"
            else:
                self.text += chr(10) + f"            '{key}': self.__{key},"

        self.text += self.text[len(self.text):len(self.text)-1] + chr(10) + '        }' + chr(10) + chr(10)
        self.text += '        par = json.dumps(par)' + chr(10) + chr(10)
        self.text += "        f = open(getcwd() + '\\parameters.json', 'w', encoding='utf-8')" + chr(10)
        self.text += '        f.write(par)' + chr(10)
        self.text += '        f.close' + chr(10)
        self.text += '# -------------------------------------------------------' + chr(10)
        self.text += "    def readParameters(self, file=getcwd() + '\\parameters.json'):" + chr(10)
        self.text += "        f = open(file, 'r', encoding='utf-8')" + chr(10)
        self.text += '        par = json.load(f)' + chr(10)
        self.text += '        f.close()' + chr(10)

        for key, item in self.dictParameters.items():
            if item[0] == 'c':
                self.text += chr(10) + f"        self.__{key} = par['{key}'].encode('ascii')"
            else:
                self.text += chr(10) + f"        self.__{key} = par['{key}']"

    def mountModuleWriteParameters(self):
        self.file = getcwd() + '\\writeParameters.py'

        self.text = 'from modulos.parameters import Parameters' + chr(10) + chr(10)
        self.text += f"p = Parameters({self.key})" + chr(10)

        for key, item in self.dictParameters.items():
            if '\\' in item[1]:
                self.text += f"p.{key} = r'{item[1]}'" + chr(10)
            else:
                self.text += f"p.{key} = '{item[1]}'" + chr(10)

        self.text += 'p.writeParametersFile()' + chr(10)

    def writeInFile(self):
        with open(self.file, 'w') as f:
            f.write(self.text)
            f.close()
