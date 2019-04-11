# coding=utf-8
import xlrd
import datetime
ctype={0:"empty",1:"string",2:"number",3:"date",4:"boolean",5:"error"}

class ExcelManage(object):
    def __init__(self, filePath):
        self.filePath = filePath
        self.workbook=xlrd.open_workbook(filename=self.filePath)

    def setFilePath(self, filePath):
        self.filePath = filePath
        self.workbook=xlrd.open_workbook(filename=self.filePath)

    def getFilePath(self):
        return self.filePath

    def getFile(self):
        return self.workbook
    '''
    将excel中的内容转换为字符串
    注:boolean类型，true转为"1",false转为"0"
    '''
    def getDataInString(self,book,cell):
        cell_ctype=cell.ctype
        if cell_ctype==0:
            return ""
        elif cell_ctype==1:
            return cell.value
        elif cell_ctype==2:
            num_value=cell.value
            if int(num_value)==num_value:
                return str(int(cell.value))
            else:
                return str(num_value)
        elif cell_ctype==3:
            date_value=xlrd.xldate_as_tuple(cell.value,book.datemode)
            if(date_value[3:])==(0,0,0):
                return datetime.date(*date_value[:3]).strftime("%Y-%m-%d")
            else:
                return datetime.datetime(*date_value).strftime("%Y-%m-%d %H:%M:%S")
        elif cell_ctype==4:
            return str(cell.value)
        elif cell_ctype==5:
            return ""
        else:
            return str(cell.value)

    '''
    读取某一sheet页,某行数据
    参数：sheet页名称,行索引(从0开始),列开始索引(从0开始),数据长度
    '''
    def getDatasByRow(self, sheetName, rowNum, sNum):
        book = self.workbook
        #sheetName=sheetName.decode('utf-8')
        try:
            sheet = book.sheet_by_name(sheetName)
        except:
            print "no sheet in %s named %s" % (self.filePath, sheetName)
            return
        dataList=[]
        maxLen=sheet.row(rowNum).__len__()
        for i in range(sNum,maxLen):
            #if(i<=maxLen-1):
                cellOri=sheet.cell(rowNum,i)
                if(cellOri.ctype==5):
                    print "cell(%d,%d) in sheet(%s) in file(%s) is error" % rowNum,i,sheetName,self.filePath
                    dataList.append("")
                else:
                    dataList.append(self.getDataInString(book,cellOri))
            #else:
                #dataList.append("")
        return dataList

    '''
    读取某一sheet页,某列数据
    参数：sheet页名称(从0开始),列索引(从0开始),行开始索引(从0开始),数据长度
    '''
    def getDatasByCol(self,sheetName,colNum,sNum):
        book = self.workbook
        try:
            sheet = book.sheet_by_name(sheetName)
        except:
            print "no sheet in %s named %s" % (self.filePath, sheetName)
            return
        dataList=[]
        maxLen=sheet.col(colNum).__len__()
        for i in range(sNum,maxLen):
            #if(i<=maxLen-1):
                cellOri=sheet.cell(i,colNum)
                if(cellOri.ctype==5):
                    print "cell(%d,%d) in sheet(%s) in file(%s) is error" % (i,colNum,sheetName,self.filePath)
                    dataList.append("")
                else:
                    dataList.append(self.getDataInString(book,cellOri))
            #else:
                #dataList.append("")
        return dataList

    '''
    从某列中找到第一个满足的字符串，返回列序号
    '''
    def getIndexByCol(self,sheetName,colNum,dist):
        book = self.workbook
        try:
            sheet = book.sheet_by_name(sheetName)
        except:
            print "no sheet in %s named %s" % (self.filePath, sheetName)
            return 9999  #9999表示有误或没有找到
        size=sheet.col(colNum).__len__()
        for i in range(0,size):
            cellStr=self.getDataInString(book,sheet.cell(i,colNum))
            if(cellStr==dist):
                return i
            else:
                continue
        return 9999 #9999表示有误或没有找到


    '''
    从某行中找到第一个满足的字符串，返回列序号
    '''
    def getIndexByRow(self,sheetName,rowNum,dist):
        book = self.workbook
        try:
            sheet = book.sheet_by_name(sheetName)
        except:
            print "no sheet in %s named %s" % (self.filePath, sheetName)
            return 9999  #9999表示有误或没有找到
        size=sheet.row(rowNum).__len__()
        for i in range(0,size):
            cellStr=self.getDataInString(book,sheet.cell(rowNum,i))
            if(cellStr==dist):
                return i
            else:
                continue
        return 9999 #9999表示有误或没有找到
    def closeExcel(self):
        self.workbook=None
if "__main__" == __name__:
    file=u"C:\\Users\\weinan\\Desktop\\TestCases.xlsx"
    print file
    em=ExcelManage(file)
    print em.getDatasByCol("TestCases",1,0)
    print (em.getDatasByCol("TestCases",1,0)[0]=="login")
    print em.getDatasByRow("case",1,0)
    print em.getDatasByRow("case",1,1)