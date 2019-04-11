# coding=utf-8
import unittest
from until.browerMangage import BrowserManage
from until.autoFormAdv import AutoForm
from lib import HTMLTestRunner,sendMail
import os
import sys
import ConfigParser
'''
配置用例文件名list及浏览器配置
'''
TESTFILE_LIST = ["testcase.xlsx" ]
data_file = "config/conf.ini"
data = ConfigParser.RawConfigParser()
data.read(data_file)
BROSWER = data.get("webservice","browser")
PROXY_IP = data.get("webservice","proxy_ip")
PROXY_PORT =  data.get("webservice","proxy_port")

'''
根据相对路径，获取绝对路径
'''
THIS_DIR = os.path.dirname(__file__).decode('gbk')  # 返回脚本路径
print THIS_DIR
LOGS_DIR = "report"
CASES_DIR = "TestCases"



def getAbsPath(relativePath):
    # 拼接路径，拼接THIS_DIR+relativePath
    return os.path.join(THIS_DIR, relativePath)


class RunTest(unittest.TestCase):
    # 类方法，cls为当前的类名
    @classmethod
    def setUpClass(cls):
        cls.bm = BrowserManage(BROSWER, PROXY_IP, PROXY_PORT)
        cls.af = AutoForm(cls.bm)

    def actions(self, arg1, arg2):
        # 获取元素的方式
        filePath = getAbsPath(CASES_DIR + "\\" + "testcase.xlsx")
        caseName = arg2
        logDir = getAbsPath(LOGS_DIR)
        res = self.af.runCase(filePath, caseName,logDir)
        state = True
        info = ""
        for i in range(0, res.__len__()):
            state = state and res.__getitem__(i).get("state")
            info = info + res.__getitem__(i).get("title") + ":" + \
                   res.__getitem__(i).get("info")
        print u"用例:" + arg1 + " " + u"用例方法:" + caseName
        print info
        self.assertEqual(True, state)
    # 闭包函数
    @staticmethod
    def getTestFunc(arg1,arg2):
        def func(self):
            self.actions(arg1,arg2)
        return func

    @classmethod
    def tearDownClass(cls):
        cls.af.quitDriver()


# 获取TestSuite
def getTestSuite(*classes):
    valid_types = (unittest.TestSuite, unittest.TestCase)
    suite = unittest.TestSuite()
    for cls in classes:
        if isinstance(cls, str):
            if cls in sys.modules:
                suite.addTest(unittest.findTestCases(sys.modules[cls]))
            else:
                raise ValueError("str arguments must be keys in sys.modules")
        elif isinstance(cls, valid_types):
            suite.addTest(cls)
        else:
            suite.addTest(unittest.makeSuite(cls))
    return suite


# 读取用例
def __generateTestCases():
    TESTCASES_LIST = []
    index = 0
    for i in range(TESTFILE_LIST.__len__()):
        t_fileName = TESTFILE_LIST.__getitem__(i)
        t_filePath = getAbsPath(CASES_DIR + "\\" + t_fileName)
        # 获取可运行状态测试集
        t_caseList = AutoForm.getTestSuiteFromStdExcel(t_filePath)
        for j in range(t_caseList.__len__()):
            testinfro = t_caseList.__getitem__(j)
            TESTCASES_LIST.append((str(index), testinfro.get("caseinfro"),testinfro.get("testmethod")))
            index = index + 1
    arglists = TESTCASES_LIST
    print arglists
    for args in arglists:
        fun = RunTest.getTestFunc(args[1], args[2])
        setattr(RunTest, 'test_func_%s' % (args[2]),fun)


# start
__generateTestCases()
if __name__ == "__main__":
    testSuite = getTestSuite(RunTest)
    print testSuite
    reportFile = "D:\\project\\weinanAuto\\report\\result1.html"
    print reportFile
    fp = file(reportFile, "wb")
    runner = HTMLTestRunner.HTMLTestRunner(stream=fp, title=u'测试报告', description=u'用例执行情况')
    runner.run(testSuite)
    #发送邮件报告
    #sendMail.sendemali(reportFile)

