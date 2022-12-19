# pip install pywin32

from win32com.client import GetObject
from subprocess import Popen
import time
from os.path import exists
from modulos.formPw import FormPw

class Sap:
    def __init__(self, sapEnv, userId=None, userPW=None, language='EN', connectByDescr = False, sapFile = ''):
        self.sapEnv = sapEnv
        self.userId = userId
        self.userPW = userPW
        self.language = language
        self.connectByDescr = connectByDescr
        self.sapFile = sapFile
        self.session = self.getSapSession()

    def getSapGui(self):
        try:
            return GetObject('SAPGUI').GetScriptingEngine
        except:
            time.sleep(0.1)
            return self.getSapGui()

    def getSapSession(self):
        if not exists(self.sapFile): self.sapFile = r'C:\Program Files (x86)\SAP\FrontEnd\SapGui\saplogon.exe'
        if not exists(self.sapFile): self.sapFile.replace(' (x86)', '')
        if not exists(self.sapFile): raise Exception('SAP Gui (saplogon.exe) not founded')

        Popen(self.sapFile)
        sapGui = self.getSapGui()
        session = None
        
        if self.connectByDescr:
            if sapGui.Connections.Count > 0:
                for conn in sapGui.Connections:
                    if conn.Description == self.sapEnv:
                        session = conn.Sessions(0)
                        break
            if session is None: session = sapGui.OpenConnection(self.sapEnv).Sessions(0)
        else:
            if sapGui.Connections.Count > 0:
                for conn in sapGui.Connections:
                    if self.sapEnv in conn.ConnectionString:
                        session = conn.Sessions(0)
                        break
            if session is None: session = sapGui.OpenConnectionByConnectionString(self.sapEnv).Sessions(0)

        if session.findById('wnd[0]/sbar').text.startswith('SNC logon'):
            session.findById('wnd[0]/usr/txtRSYST-LANGU').text = self.language
            session.findById('wnd[0]').sendVKey(0)
            session.findById('wnd[0]').sendVKey(0)
        elif session.Info.User == '':
            if not self.userPw:
                fPW = FormPw()
                self.userId = fPW.userId
                self.userPw = fPW.userPw
            session.findById('wnd[0]/usr/txtRSYST-BNAME').text = self.userId
            session.findById('wnd[0]/usr/pwdRSYST-BCODE').text = self.userPW
            session.findById('wnd[0]/usr/txtRSYST-LANGU').text = self.language
            session.findById('wnd[0]').sendVKey(0)
        
        while self.fieldExists('wnd[1]'): session.findById('wnd[1]').close()

        session.findById('wnd[0]').maximize()
        return session

    def executeTransaction(self, transacao):
        '''
        Esta rotina irá colocar o '/n' na frente da transação informada e dará o enter
        '''
        self.session.findById('wnd[0]/tbar[0]/okcd').text = '/n' + transacao
        self.session.findById('wnd[0]').sendVKey(0)

    def exportTxtFile(self, pathName, fileName, button=11):
        '''
        Esta rotina irá acionar os comandos de finalização de exportação de TXT no SAP
        buttons: 0 - Gerar | 11 - Substuir | 7 - Ampliar
        '''
        self.session.findById('wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]').select() #não converter
        self.session.findById('wnd[1]/tbar[0]/btn[0]').press() #confirmar
        self.session.findById('wnd[1]/usr/ctxtDY_PATH').text = pathName
        self.session.findById('wnd[1]/usr/ctxtDY_FILENAME').text = fileName
        self.session.findById(f'wnd[1]/tbar[0]/btn[{button}]').press()        

    def fieldExists(self, field):
        try:
            var = self.session.findById(field)
            return var
        except:
            return False

    def getStatusBar(self):
        return self.session.findById('wnd[0]/sbar').Text

    def removeStatusBar(self, sendEnter=False):
        if sendEnter: self.session.findById('wnd[0]').sendVKey(0)
        while self.session.findById('wnd[0]/sbar').Text != '' and self.session.findById('wnd[0]/sbar').MessageType != 'S':
            self.session.findById('wnd[0]').sendVKey(0)
            if self.session.findById('wnd[0]/sbar').MessageType == 'E': return self.session.findById('wnd[0]/sbar').Text
        
        if self.fieldExists('wnd[1]'):
            iRow = 3
            statusError = ''
            try:
                while True:
                    if self.session.findById(f'wnd[1]/usr/lbl[3,{iRow}]').IconName == 'S_LEDR':
                        statusError += self.session.findById(f'wnd[1]/usr/lbl[7,{iRow}]').text
                    iRow += 1
            finally:
                self.session.findById('wnd[1]/tbar[0]/btn[0]').press()
                self.session.findById('wnd[0]/tbar[1]/btn[6]').press()
                self.session.findById('wnd[1]/usr/btnSPOP-OPTION2').press()
                return statusError

    def multipleSelection(self, fieldFullId):
        '''
        Copiar dados na área de transferência do Windows antes de chamar este método
        '''
        self.session.findById(fieldFullId).press() # Seleção Múltipla
        self.session.findById('wnd[1]/tbar[0]/btn[16]').press() # Apagar tudo
        self.session.findById('wnd[1]/tbar[0]/btn[24]').press() # Colar
        self.session.findById('wnd[1]/tbar[0]/btn[8]').press() # Transferir           

    def clearMultipleSelection(self, fieldFullId):
        self.session.findById(fieldFullId).press() # Seleção Múltipla
        self.session.findById('wnd[1]/tbar[0]/btn[16]').press() # Apagar tudo
        self.session.findById('wnd[1]/tbar[0]/btn[8]').press() # Transferir  

