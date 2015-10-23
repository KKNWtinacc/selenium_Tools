# -*- coding: utf-8 -*-
'''
对读取文件的相关处理

@author: wanglong
'''
import os
import csv

class File_method(object):

    '''
    classdocs
    '''


    def get_file_path(self,file_name):
        '''
                返回文件全路径   
        ep:
        obj = File_method.File_method()
        file_path=obj.get_file_path("rrr.csv")
        '''
        file_path = os.getcwd()
        path=  file_path[0:len(file_path)-3]
        all_file_path = path+"res\\"+file_name
        return all_file_path
                  
        
    
    def get_file_value(self,file_path):
        '''
                        返回文件内容,支持txt和csv
         eq     obj = File_method()   
                file_path=obj.get_file_path(u"登录.csv")
                all_values = obj.get_file_value(file_path)         
        '''
        name_last_3 = file_path[-3:]
        if name_last_3.lower() == "txt":
            user_file = open(file_path,'r')
            values = user_file.readlines()
            user_file.close()
        elif name_last_3.lower()== "csv":
            values = csv.reader(file(file_path,"rb"))
        else:
            print("文件错误")  
       
        return values
    
    def line_to_list(self,line_value):
        '''
                    将文本文件行内容转换为数组返回
        '''
        list_value = line_value.split(",")
        
        return list_value
    
    def csv_list(self,file_name,get_column):
        """
                将CSV文件的指定列返回成list
        eq csv_list(u"登录.csv", 0)
        """
        get_list = []
        obj = File_method()   
        file_path=obj.get_file_path(u"登录.csv")
        all_values = obj.get_file_value(file_path)
        for id_list in all_values:
            get_list.append(id_list[get_column])
        return get_list
    
    def csv_dict(self,file_name,key_colum,value_colum): 
        """
                将CSV文件的指定两列转换成字典
        eq: csv_dict(u"登录.csv",0,1)
        """   
        get_dict = {}
        obj = File_method()
        file_path=obj.get_file_path(file_name)
        all_values = obj.get_file_value(file_path)
        for id_name in all_values:
            get_dict.setdefault(id_name[key_colum],id_name[value_colum])
        return get_dict
        
    def get_dir_list(self,dir_path):
        '''
                        获得制定目录下的所有文件
        :param dir_path:路径
        '''  
        try:
            return os.listdir(dir_path)
        except (NameError,WindowsError):
            print ("文件名称错误")