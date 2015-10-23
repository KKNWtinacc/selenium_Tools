#coding= utf-8
'''
Created on 20150921

@author: wanglong
'''
import File_method
import xlrd
class Excel_rd(object):
    '''
            对excel文件的相关操作
    '''


    def __init__(self):
       self._file = File_method.File_method()
#        self._xl_name = u"登录.xlsx"
        
    def get_excel_table(self,file_name,excel_sheet):
        '''
                            获得excel指定sheet数据
        :param file_name: 文件名
        :param excel_sheet:表名
        eq:
            obj = Excel_rd.Excel_rd()
            _table = obj.get_excel_table(u"登录.xlsx", u"Sheet1") 
        '''
       
        try:
            _file_path = self._file.get_file_path(file_name)
            
            _data = xlrd.open_workbook(_file_path)
        except (IOError):
            print "文件不存在 "+file_name
            
        try:
            _table = _data.sheet_by_name(excel_sheet)
            return _table
        except (UnboundLocalError):
            print "sheet页不存在 "+excel_sheet
        
        
    def get_row_value(self,excel_table,number):  
        '''
                        获得excel制定sheet的指定行数据，从0开始
        eq:
            obj = Excel_rd.Excel_rd()
            _table = obj.get_excel_data(u"登录.xlsx", u"Sheet1")
            _row = obj.get_row_value(_table, 0) 
            
        :param excel_table:获取的表数据
        :param number:指定的行数。 从0开始
        '''
       
        return excel_table.row_values(number)
    
    def get_col_value(self,excel_table,number):
        '''
                         获得excel制定sheet指定列数据，从0开始
        :param excel_table:获取的表内容
        :param number:指定的行数
        '''
        
        return excel_table.col_values(number)
        
    def get_row_number(self,excel_table):  
        '''
                         返回excel的行数
        :param excel_table:获得表的内容
        '''
       
        return excel_table.nrows
    
    def get_col_number(self,excel_table):  
        '''
                         返回excel的列数
        :param excel_table: 获得的表信息
        '''
      
        try:
            return excel_table.ncols  
        except AttributeError:
            print "表内容错误"
    
    def get_cell_value(self,excel_table,rowx,colx): 
        '''
                        获得制定单元格数据
        :param excel_table:获得的表数据
        :param rowx:行坐标
        :param colx:列坐标
        eq:
            obj = Excel_rd.Excel_rd()
            _table = obj.get_excel_table(u"登录.xlsx", u"Sheet1") 
            cell_value = obj.get_cell_value(_table, 1, 1)
        '''
        
        return excel_table.cell(rowx, colx).value
    
    def get_cellname(self,rowx,colx):
        '''
                            输入坐标，返回单元格名称 (5, 7) => 'H6' 
        :param rowx: 行坐标
        :param colx: 列坐标
        '''
        
        return xlrd.cellname(rowx, colx)
        