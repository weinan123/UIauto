# coding=utf-8
import datetime,os
from operate import Operate
from excelManage import ExcelManage
from selenium.webdriver.common.by import By
from elements import getElement

def getErrInfo(xpath, op, value, errInfo):
    return "Xpath(%s),Operate(%s),Value(%s):%s" % (xpath, op, value, errInfo)


class AutoForm(object):
    def __init__(self, bm):
        self.YES = "yes"
        self.NO = "no"
        self.SPLITMARK = "###"
        self.XPATHARG = "{@@}"
        self.driver = bm.getDriver()
        self.proxyIp = bm.getIp()
        self.proxyPort = bm.getPort()
        self.operate = Operate(self.driver)
        self.currentWindow = self.driver.current_window_handle
        self.defaultWindow = self.driver.current_window_handle

    def getCurrentWindow(self):
        return self.currentWindow

    def getDefaultWindow(self):
        return self.defaultWindow

    def setCurrentWindow(self, cWindow):
        self.currentWindow = cWindow

    def setDefaultWindow(self, dWindow):
        self.defaultWindow = dWindow

    def getDriver(self):
        return self.driver

    '''
    结合operate中操作，生成带窗口状态处理的新操作
    '''

    def doOperate(self,by,xpath,value,op):
        operateDic = {
            "closeotherwindows": self.closeOtherWindows,
            "switchtowindowbyclick": self.switchToWindowByClick,
            "switchtodefaultwindow": self.switchToDefaultWindow,
            "setcurrwinasdefault": self.setCurrWinAsDefault
        }
        try:
            if (operateDic.has_key(op.lower())):
                return operateDic.get(op.lower())(xpath,value,by)
            else:
                # 处理非WebDriver处理的请求方法
                if (op.lower() == "sameaslocalfile"):
                    if (self.proxyIp <> "" and self.proxyIp <> None):
                        t_proxy = self.proxyIp + str(self.proxyPort)
                    else:
                        t_proxy = ""
                    return self.operate.sameAsLocalFile(by,xpath, value, proxy=t_proxy)
                # 处理WebDriver处理的请求方法
                else:
                    return self.operate.operate(by, xpath, value,op.lower())
        except  Exception as e:
            return {"state": False, "info": getErrInfo(xpath, op, value, str(e))}

    '''
    设置当前窗口作为默认窗口
    '''

    def setCurrWinAsDefault(self, xpath, value):
        result = {}
        try:
            self.setDefaultWindow(self.driver.current_window_handle)
            self.setCurrentWindow(self.driver.current_window_handle)
            result = {"state": True, "info": ""}
        except Exception as e:
            result = {"state": False, "info": getErrInfo(xpath, "setCurrWinAsDefault", value, str(e))}
        return result

    '''
    关闭当前窗口以外的窗口
    '''

    def closeOtherWindows(self, xpath, value):
        result = {}
        try:
            allWindows = self.driver.window_handles
            for i in range(0, allWindows.__len__()):
                tempWindow = allWindows.__getitem__(i)
                if (self.currentWindow <> tempWindow):
                    self.driver.switch_to.window(tempWindow)
                    self.driver.close()
                self.driver.switch_to.window(self.currentWindow)
                self.setDefaultWindow(self.driver.current_window_handle)
                self.setCurrentWindow(self.driver.current_window_handle)
                result = {"state": True, "info": ""}
        except Exception as e:
            result = {"state": False, "info": getErrInfo(xpath, "closeOtherWindows", value, str(e))}
        return result

    '''
    点击某个元素后，打开新窗口，并指向新窗口
    '''

    def switchToWindowByClick(self, xpath, value,by):
        result = {}
        try:
            oriWindows = self.driver.window_handles
            self.driver.find_element(by,xpath).click()
            self.driver.implicitly_wait(3)
            nowWindows = self.driver.window_handles
            newWindow = None
            for i in range(0, nowWindows.__len__()):
                t = nowWindows.__getitem__(i)
                try:
                    oriWindows.index(t)
                    continue
                except:
                    newWindow = t
                    break
            if (newWindow <> None):
                self.driver.switch_to.window(newWindow)
                self.setCurrentWindow(self.driver.current_window_handle)
                result = {"state": True, "info": ""}
            else:
                result = {"state": False, "info": getErrInfo(xpath, "switchToWindowByClick", value,
                                                             "can't find new Window by click this element!")}
        except Exception as e:
            result = {"state": False, "info": getErrInfo(xpath, "switchToWindowByClick", value, str(e))}
        return result

    '''
    指向原窗口
    '''

    def switchToDefaultWindow(self, xpath, value):
        result = {}
        try:
            self.driver.switch_to.window(self.getDefaultWindow())
            self.setCurrentWindow(self.driver.current_window_handle)
            result = {"state": True, "info": ""}
        except Exception as e:
            result = {"state": False, "info": getErrInfo(xpath, "switchToDefaultWindow", value, str(e))}
        return result
    '''
    合并way数组，xpath数组，操作数组，值数组.
    构成字典列表
    '''

    def mergeInputDatas(self,filePath ):
        em = ExcelManage(filePath)
        casename_array=em.getDatasByCol('case',0,1)
        caseinfro_array = em.getDatasByCol('case',1,1)
        way_array=em.getDatasByCol('case',2,1)
        xpath_array=em.getDatasByCol('case',3,1)
        option_array=em.getDatasByCol('case',4,1)
        data_array=em.getDatasByCol('case',5,1)
        length = casename_array.__len__()
        inputDatas = []
        for i in range(0, length):
            if (i >= caseinfro_array.__len__()):
                t_caseinfro = ""
            else:
                t_caseinfro = caseinfro_array[i]
            if (i >= way_array.__len__()):
                t_way = ""
            else:
                t_way = way_array[i]

            if (i >= xpath_array.__len__()):
                t_xpath = ""
            else:
                t_xpath = xpath_array[i]

            if (i >= option_array.__len__()):
                t_op = ""
            else:
                t_op = option_array[i]

            if (i >= data_array.__len__()):
                t_value = ""
            else:
                t_value = data_array[i]
            inputDatas.append({"casename":casename_array[i],"caseinfro":t_caseinfro,"way":t_way,"xpath": t_xpath, "op": t_op, "value": t_value})
        return inputDatas

    def fixTheForm(self,caseinfro,by,xpath,value,op ,logDir=None):
        result = {"state": True, "info": self.getNowStrftime(), "title": ">>>" + caseinfro}
        t_result = self.doOperate(by,xpath,value,op)
        if (not t_result.get("state")):
            if ((xpath or "").__len__() > 0 and getElement.isElementPresent(self.driver, by, xpath)):
                getElement.scrollToElement(self.driver, self.driver.find_element(by,xpath))

            result = self.mergeResult(result, t_result)
        if (result.get("state")):
            result.__setitem__("info", result.get("info") + " PASS\n")
        return result

    '''
    获取当前时间的格式化字符串
    '''

    def getNowStrftime(self):
        return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    def getNowStrftime2(self):
        return datetime.datetime.now().strftime("%Y%m%d%H%M%S")

    '''
    合并各个步骤的测试结果(没用)
    '''

    def mergeResult(self, result, t_result):
        result.__setitem__("state", result.get("state") and t_result.get("state"))
        result.__setitem__("info", result.get("info") + "\n" + t_result.get("info"))
        return result
    '''
    获取查找元素的方式
    '''
    def getWays(self,way):
        by=''
        if way=='id':
            by= By.ID
        elif way=="name":
            by=By.NAME
        elif way == "xpath":
            by = By.XPATH
        elif way == "css":
            by = By.CSS_SELECTOR
        elif way == "text":
            by = By.LINK_TEXT
        elif way == "classname":
            by = By.CLASS_NAME
        return by

    '''
    运行指定格式sheet页case中的测试用例
    '''

    def runCase(self, filePath,casename,logsDir=None):
        #获取case中可执行的用例数
        resultList = []
        datas = self.mergeInputDatas(filePath)
        for i in range(0,datas.__len__()):
            tmp = datas.__getitem__(i)
            runcase = tmp.get('casename')
            if runcase==casename and runcase !="":
                way = tmp.get("way")
                caseinfro=tmp.get("caseinfro")
                xpath=tmp.get("xpath")
                op=tmp.get("op")
                value=tmp.get('value')
                by=self.getWays(way)
                t_result = self.fixTheForm(caseinfro,by,xpath,value,op ,logDir=logsDir)
                resultList.append(t_result)
        return resultList

    '''
    获取testcase指定格式excel中的测试集，可运行状态是否
    '''

    @staticmethod
    def getTestSuiteFromStdExcel(filePath):
        em = ExcelManage(filePath)
        caseArray = em.getDatasByCol('testcase', 0, 1)
        stateArray = em.getDatasByCol('testcase', 1, 1)
        caseinfro = em.getDatasByCol('testcase', 2, 1)
        print caseinfro
        print stateArray.__len__()
        runCaseList = []
        for i in range(0, stateArray.__len__()):
            if (stateArray[i] == 'Y'):
                runCaseList.append({"caseinfro":caseinfro[i],"testmethod":caseArray[i]})
        return runCaseList

    '''
    退出WebDriver
    '''

    def quitDriver(self):
        self.driver.quit()




