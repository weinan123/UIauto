#coding=utf-8
import os
from selenium import webdriver
from selenium.webdriver.common.proxy import *
class BrowserManage(object):
    def __init__(self,browser="",ip="",port=0,timeout=30):
        self.browser=browser
        self.ip = ip
        self.port = port
        self.timeout = timeout
    def setBrowser(self, browser):
        self.browser = browser
    def setProxy(self, ip, port):
        self.ip = ip
        self.port = port
    def setTimeout(self, timeout):
        self.timeout = timeout
    def getBrowser(self):
        return self.browser
    def getIp(self):
        return self.ip
    def getPort(self):
        return self.port
    def getDriver(self):
        return BrowserManage.startDriver(_browser=self.browser,_ip=self.ip,_port=self.port,_timeout=self.timeout)

    #获取对应的webdriver
    @staticmethod
    def startDriver(_browser="chrome",_ip="",_port=0,_timeout=30):
        driver=None
        browser = _browser.lower()
        ip=_ip
        port=_port
        timeout=_timeout
        #默认chrome浏览器
        if not("firefox","chrome","ie").__contains__(browser):
            browser="chrome"
        myProxy = ip + ":" + str(port)
        if browser == "firefox":
            if (ip <> "" and ip <> 0):
                proxy = Proxy({
                    'proxyType': ProxyType.MANUAL,
                    'httpProxy': myProxy,
                    'ftpProxy': myProxy,
                    'sslProxy': myProxy,
                    'noProxy': ''
                })
                driver = webdriver.Firefox(proxy=proxy)
            else:
                driver = webdriver.Firefox()
        elif browser == "chrome":
            if (ip <> "" and ip <> 0):
                chrome_options = webdriver.ChromeOptions()
                chrome_options.add_argument('--proxy-server=%s' % myProxy)
                driver = webdriver.Chrome(chrome_options=chrome_options)
            else:
                cmd_switchProxy="reg add \"HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Internet Settings\" /v ProxyEnable /t REG_DWORD /d 0 /f & "+\
                        "reg add \"HKCU\\Software\\Microsoft\Windows\\CurrentVersion\\Internet Settings\" /v ProxyServer /d \"\" /f & "+\
                        "ipconfig /flushdns"
                os.system(cmd_switchProxy)
                chromedriver = "C:\Python27\Scripts\chromedriver.exe"
                os.environ["webdriver.chrome.driver"] = chromedriver
                driver = webdriver.Chrome(chromedriver)
                #driver = webdriver.Chrome()

        elif browser == "ie":
              #ie无法手动配置代理，配置系统代理
              if (ip <> "" and ip <> 0):
                    cmd_switchProxy="reg add \"HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Internet Settings\" /v ProxyEnable /t REG_DWORD /d 1 /f & "+\
                        "reg add \"HKCU\\Software\\Microsoft\Windows\\CurrentVersion\\Internet Settings\" /v ProxyServer /d \""+myProxy+"\" /f & "+\
                        "ipconfig /flushdns"
                    os.system(cmd_switchProxy)
              else:
                    cmd_switchProxy="reg add \"HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Internet Settings\" /v ProxyEnable /t REG_DWORD /d 0 /f & "+\
                        "reg add \"HKCU\\Software\\Microsoft\Windows\\CurrentVersion\\Internet Settings\" /v ProxyServer /d \"\" /f & "+\
                        "ipconfig /flushdns"
                    os.system(cmd_switchProxy)
              driver = webdriver.Ie()
        driver.set_page_load_timeout(timeout)
        driver.accept_next_alert = True
        return driver