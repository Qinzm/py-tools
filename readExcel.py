# coding=utf-8
import xlrd,os
class ReadExcel(object):
    """
    excel数据读取和数据解析类
    """
    def openExcel(self,fileName):
        """
                读取excel
        """
        try:
            data = xlrd.open_workbook(fileName)
            return data
        except Exception,e:
            print str(e)

    def excelTableByName(self,fileName,colnameindex=0,byName=u'sheet1'):
        """
        #根据名称获取Excel表格中的数据
        fileName:Excel文件路径     
        colnameindex:头列名所在行
        byName:名称
        """
        data = self.openExcel(fileName)

        table = data.sheet_by_name(byName)
        nrows = table.nrows #行数
        listData=[]
        colnames =  table.row_values(colnameindex)
        for rownum in range(1,nrows):
            row = table.row_values(rownum)
            if row:
                app = {}
                for i in range(len(colnames)):
                    app[colnames[i]] = row[i]
                listData.append(app)
        return colnames,listData