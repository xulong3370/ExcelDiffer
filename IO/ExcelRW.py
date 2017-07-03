#coding=utf-8
"""
读取和写入excel的模块
"""
import xlrd
import datetime
CellType = {0:"empty",1:"string", 2:"number", 3:"date", 4:"boolean", 5:"error"}
class ExtractExcel(object):
    """
    解析操作Excel
    :param  path:excel 文件路径
    """
    def __init__(self,path):
        self.path = path.decode('utf-8')
        self.workbook = xlrd.open_workbook(self.path)
    def getCellData(self,sheet,row,col):
        """
        获得单元格数据
        :param sheet:选择的sheet表格名字
        :param row:行 int
        :param col:列 int
        :return : 单元格数据 string int float date
        """
        if(isinstance(sheet,str)):
            value = self.workbook.sheet_by_name(sheet.decode("utf8")).cell(row,col).value
        elif(isinstance(sheet,unicode)):
            value = self.workbook.sheet_by_name(sheet).cell(row, col).value
        else:
            value = self.workbook.sheets()[sheet].cell(row,col).value
        return value
    def getCellType(self,sheet,row,col):
        """
        获得单元格数据类型
        :param sheet: sheet 索引 int
        :param row: 行 int
        :param col: 列 int
        :return: 单元格类型 string
        """
        if (isinstance(sheet, str)):
            typeIndex = self.workbook.sheet_by_name(sheet.decode("utf8")).ctype
        elif(isinstance(sheet,unicode)):
            typeIndex = self.workbook.sheet_by_name(sheet).ctype
        else:
            typeIndex = self.workbook.sheets()[sheet].cell(row,col).ctype
        return CellType[typeIndex]
    def getMergeCellInfo(self,sheet):
        """
        获得sheet合并的单元格信息
        :param sheet: sheet 索引 int
        :return: [row,rowRange,col,colRange]合并单元格信息
        """
        if(isinstance(sheet,str)):
            mergeCell = self.workbook.sheet_by_name(sheet.decode("utf8")).merged_cells
        elif(isinstance(sheet,unicode)):
            mergeCell = self.workbook.sheet_by_name(sheet).merged_cells
        else:
            mergeCell = self.workbook.sheets()[sheet].merged_cells
        return mergeCell
    def getSheetMaxSize(self,sheet):
        """
        获得表格最大尺寸
        :param sheet:  表格 索引
        :return: 返回sheet大小
        """
        if(isinstance(sheet,str)):
            return self.workbook.sheet_by_name(sheet.decode("utf8")).nrows,\
                   self.workbook.sheet_by_name(sheet.decode("utf8")).ncols
        elif(isinstance(sheet,unicode)):
            return self.workbook.sheet_by_name(sheet).nrows,self.workbook.sheet_by_name(sheet).ncols
        else:
            return self.workbook.sheets()[sheet].nrows, self.workbook.sheets()[sheet].ncols
    def getSheetName(self,sheet):
        """
        :param sheet: 表格索引
        :return: 返回sheet名字
        """
        return self.workbook.sheets()[sheet].name
    def getExcelSheetsNameList(self):
        """
        :return: 返回当前excel中所有sheet name的链表
        """
        sheetList = []
        for sheet in self.workbook.sheets():
            sheetList.append(sheet.name)
        return sheetList
    def dateFormat(self,date,format):
        """
        格式化Excel单元格
        :param date: 时间
        :param format: 返回的时间格式
        :return: 格式化时间
        """
        year, month, day, hour, minute, second = xlrd.xldate_as_tuple(date,self.workbook.datemode)
        dateFormat = datetime.datetime(year, month, day, hour, minute, second)
        return dateFormat.strftime(format)
    def getFilePath(self):
        return self.path







