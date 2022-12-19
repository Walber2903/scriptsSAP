# Instalações via pip:
# pip install requests
# pip install beautifulsoup4
# pip install selenium
# pip install pywin32

import requests
from bs4 import BeautifulSoup
from enum import Enum
from selenium.webdriver.remote import webelement
from os import mkdir, path
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from os import getcwd
from os import getenv as env, remove, mkdir
from os.path import exists
from win32com.client import Dispatch
from selenium import webdriver as wd
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.chrome.options import Options as chromeOpt
from zipfile import ZipFile as zip
import re
import time

class browsers(Enum):
    chrome = ('chromedriver.exe', 'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe')
    edge = ('msedgedriver.exe', 'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe')
    ie = ('iedriver.exe', 'C:\Program Files\Internet Explorer\iexplore.exe')
    firefox = ('geckodriver.exe', 'C:\Program Files\Mozilla Firefox\firefox.exe')

class Webscrapping:
    from selenium.webdriver.remote.webelement import WebElement as webElem

    def __init__(self, browser, visible=True, url='', timeout=0):
        path = f'{env("localappdata")}\\drivers\\'
        if not exists(path): mkdir(path)
        self.browser = browser.name
        self.driverExe = browser.value[0]
        self.driver = path + self.driverExe
        self.browserFullname = browser.value[1]
        self.visible = visible
        self.url = url
        self.webdriver = self.getDriver()
        if url: self.webdriver.get(url)
        if timeout>0: self.webdriver.implicitly_wait(timeout)
        self.we: webelement.WebElement = None
        self.wes = None

    def getDriver(self):
        try:
            if self.browser == 'chrome':
                opt = chromeOpt()
                opt.headless = not self.visible
                driver = wd.Chrome(self.driver, options=opt)
                driver.maximize_window()
            elif self.browser == 'edge':
                driver = wd.Edge(self.driver)
                driver.maximize_window()
            elif self.browser == 'firefox':
                driver = wd.Firefox(executable_path=self.driver)
                driver.maximize_window()
            elif self.browser == 'ie':
                driver  = wd.Ie(self.driver)
                driver.maximize_window()
        except WebDriverException as e:
            if e.msg.startswith(f"'{self.driverExe}' executable needs to be in PATH") or e.msg.startswith('session not created: This version'):
                self.downloadDriver()
                return self.getDriver()
            else:
                raise Exception(str(e))
        except Exception as e:
            raise Exception(str(e))

        return driver

    def driverURL(self):
        if self.browser == 'chrome':
            ver = self.getBrowserActualVersion()[:9]
            url = 'https://chromedriver.chromium.org/downloads'
            response = requests.get(url)
            htmlText = response.text
            soup = BeautifulSoup(htmlText, 'html.parser')
            try:
                ver = soup.find(class_='XqQF9c', text=re.compile(ver)).contents[0].split()[1]
            except:
                ver = soup.find(class_='XqQF9c').contents[0].split()[1]
            url = f'https://chromedriver.storage.googleapis.com/{ver}/chromedriver_win32.zip'
        elif self.browser == 'edge':
            ver = self.getBrowserActualVersion()
            url = f'https://msedgedriver.azureedge.net/{ver}/edgedriver_win64.zip'
        elif self.browser == 'ie':
            url = 'https://drive.google.com/u/0/uc?id=1NosL-ZdTlw3oOHgBNcPDLEfNfEEU-fwu&export=download'
        elif self.browser == 'firefox':
            url = 'https://github.com/mozilla/geckodriver/releases'
            response = requests.get(url)
            htmlText = response.text
            soup = BeautifulSoup(htmlText, 'html.parser')
            ver = soup.find(class_='Link--primary', href=re.compile('/mozilla/geckodriver/releases/tag/')).contents[0]
            url = f'https://github.com/mozilla/geckodriver/releases/download/v{ver}/geckodriver-v{ver}-win64.zip'
        return url
    
    def getBrowserActualVersion(self):
        fso = Dispatch('Scripting.FileSystemObject')
        if not exists(self.browserFullname): self.browserFullname = self.browserFullname.replace(' (x86)','')
        if not exists(self.browserFullname): raise Exception(f'Browser {self.browser} not installed')
        return fso.getFileVersion(self.browserFullname)

    def downloadDriver(self):
        url = self.driverURL()
        response = requests.get(url)
        zipFile = 'driver.zip'
        
        with open(zipFile, 'wb') as f:
            f.write(response.content)
            f.close()
        with open(self.driver, 'wb') as f:
            f.write(zip(zipFile).read(self.driverExe))
            f.close()

        remove(zipFile)

    def getElementBy(self, by: By, strElement: str, weParent: webElem = None, timeout: int = 20) -> webelement.WebElement:
        if timeout:
            try:
                if weParent:
                    self.we = weParent.find_element(by, strElement)
                else:
                    self.we = self.webdriver.find_element(by, strElement)
                return self.we
            except:
                timeout -= 1
                time.sleep(1)
                self.we = self.getElementBy(by, strElement, weParent, timeout) 
                return self.we

    def getElementsBy(self, by: By, strElement: str, weParent: webElem = None, timeout: int = 20):
        if timeout:
            try:
                if weParent:
                    self.wes = weParent.find_elements(by, strElement)
                else:
                    self.wes = self.webdriver.find_elements(by, strElement)
                return self.wes
            except:
                timeout -= 1
                time.sleep(1)
                self.wes = self.getElementsBy(by, strElement, weParent, timeout) 
                return self.wes

    def click(self, index=0):
        if index:
            self.wes[index].click()
        else:
            self.we.click()
    
    def sendKeys(self, text, index=0):
        if index:
            self.wes[index].clear()
            self.wes[index].send_keys(text)  
        else:
            self.we.clear()
            self.we.send_keys(text)

    def close(self):
        self.webdriver.close()
