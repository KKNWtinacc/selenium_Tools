# -*- coding: utf-8 -*-
'''
Created on 20150819

@author: wanglong
'''
from selenium import webdriver
import time,os,string
from base import File_method
from base import Excel_rd
import ConfigParser
from ConfigParser import NoSectionError
from selenium.common.exceptions import NoSuchElementException
# 
# reload(sys)
# sys.setdefaultencoding('utf-8')
def get_driver(profile_dir):
    '''
                返回浏览器句柄
    :param profile_dir:
    '''
    _profile_ = webdriver.FirefoxProfile(profile_dir)
    _driver_ = webdriver.Firefox(_profile_)
    _driver_.maximize_window()
    return _driver_

def is_repeat(sen,opt):
    '''
            判断配置文件的内容是否存在， 如果存在则返回对应的值。如不存在，则返回false
    :param cfg: config配置文件
    :param sen: 标题部分
    :param opt: 标题内选项
    '''
    try:
        return config.get(sen, opt)
    except ConfigParser.NoOptionError,ConfigParser.NoSectionError:
        print "未找到标题"+sen+"或内容为"+opt
        return False
    
def is_in_dir(file_name,dir_list):
    '''
            返回文件是否存在指定的目录中
    :param file_name:
    :param dir_list:
    '''
    return file_name in dir_list

def get_sention_count(_sention):
    '''
            获得制定区域的行数
    :param _sention:区域标签
    '''
    try:
        _list_sention_ = config.items(_sention)
        return len(_list_sention_)
    except NoSectionError:
        print "文件中无此字段"
        return 0
def get_str(_row_value,add_str =""): 
    '''
                            将excel整行的内容。转换成运行的语句
    :param row_value:excel行内容
    '''

   
    if _row_value[6] == '':
        _para2 = "()"
    else:
    #            
        _para2 = "(\""+_row_value[6]+str(add_str)+"\")"
    #            
        
    if _row_value[5] !='':
        _para2 = _row_value[5]+_para2  
                     
    if _row_value[3] == '':
        _cmd = "dr."+_row_value[2]+_para2
    else :
         
        _cmd = "dr."+_row_value[2]+"(\""+_row_value[3]+"\",\""+str(_row_value[4])+"\")."+_para2
    print _cmd
    return _cmd
def run_cmd(dr,_cmd_str):
    _sleep_time_ = config.getint("STEP", "sleep")

    '''
            运行指定语句
    :param dr:浏览器句柄
    :param _cmd_str:运行的字符串语句
    '''
   
    if ".sleep(" in _cmd_str:
        time.sleep(string.atof(_cmd_str[10:9+len(_cmd_str)-11]))
    else:
             
        try:
            eval(_cmd_str)
        except NoSuchElementException :
            print (u"未找到元素"+_cmd_str)
        
        time.sleep(_sleep_time_)
        
        

if __name__ == "__main__":      

    
    _file = File_method.File_method()
    _excel_file = Excel_rd.Excel_rd()
    _this_path = os.getcwd()
    _res_path =  _this_path[0:len(_this_path)-3]+"res\\"
    _path =_file.get_file_path('config.ini')
    
    config=ConfigParser.ConfigParser()
    
    
    
    try:
        config.readfp(open(_path))
      
    except IOError:
        print "文件路径错误，路径为 "+_path
    
    _dir_list = _file.get_dir_list(_res_path) #目录内文件名
    _profile_dir = is_repeat("WEB_CONFIG", "profile_dir")
    _sheet_name = is_repeat("EXCEL_SHEET", "sheet_name")
#     print _profile_dir,_sheet_name
    _driver_ = get_driver(_profile_dir)
    
    _list_sention_count_ = get_sention_count("EXCEL_FILE") 
    print    _list_sention_count_
    for i in range(_list_sention_count_):#取list里的行数遍历
        _list_sention = is_repeat( "EXCEL_FILE", "list"+str(i+1))#获得配置文件对应的值
        
        print _list_sention
        if is_in_dir(_list_sention, _dir_list):#判断配置文件内写的文件名在目录中
            _file_table_ = _excel_file.get_excel_table(_list_sention, _sheet_name)#获取表文件
            _sheet_num_ = _excel_file.get_row_number(_file_table_)#取表的行数
            _repeat_conut = 0 #需要循环的总次数
            _repeat_start_num =0 #当前的循环次数
            _for_end_row = 0 #结束的行号
            j = 1
            _repeat_flag_= False
            while j<_sheet_num_:
                _cell_value = _excel_file.get_cell_value(_file_table_, j, 2)
                if _cell_value == "for":
                    _for_start_row = j
                    if int(_repeat_conut) != 0 and _for_start_row > _for_end_row:
                        _repeat_start_num = 0 
                    _repeat_conut = _excel_file.get_cell_value(_file_table_, j, 6)
                    
                   
                    
                elif _cell_value =="forend":
                    _for_end_row = j
                    _repeat_start_num+=1
                    if _repeat_start_num < int(_repeat_conut):
                        j = _for_start_row
                        _repeat_flag_= True
                    else:
                        _repeat_flag_= False   
                else:
                    _row_value_ = _excel_file.get_row_value(_file_table_, j)
                    if _repeat_flag_ and _excel_file.get_cell_value(_file_table_, j, 7) == "Y":
                        _cmd_str_ = get_str(_row_value_,_repeat_start_num)
                    else:
                        _cmd_str_ = get_str(_row_value_)
                    run_cmd(_driver_, _cmd_str_)
                j+=1
                
        else:
            print "指定文件 "+_list_sention+" 不存在" 
            


        

