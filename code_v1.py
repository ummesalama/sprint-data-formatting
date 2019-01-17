# -*- coding: utf-8 -*-
"""
Created on Mon Jan 14 15:25:14 2019

@author: User
"""

import pandas as pd
import os

def read_csv_file(csv_file_path):
    if(os.path.exists(csv_file_path)):
        data_frame = pd.read_csv(csv_file_path, usecols=['Issue key','Summary','Status','Component/s','Labels','Log Work'])
        new_data_frame = data_frame['Log Work'].str.split(';', n = 3, expand = True) 
        data_frame['Worklog Comment']= new_data_frame[0]
        data_frame['Worklog Date']= new_data_frame[1]
        data_frame['Worked By']= new_data_frame[2]
        new_data_frame.loc[new_data_frame[3].notnull(), 3] = new_data_frame.loc[new_data_frame[3].notnull(), 3].apply(int)
        data_frame['Time Spent (hours)'] = new_data_frame[3]/3600
        data_frame.drop(columns =["Log Work"], inplace = True) 
        new_data_frame.to_csv(csv_file_path, sep='\t', encoding='utf-8')
        data_frame.rename(columns={'Labels': 'Platform'}, inplace=True)
        data_frame.rename(columns={'Log Work': 'Worklog Comment'}, inplace=True)
        data_frame = data_frame[['Issue key','Summary','Status','Component/s','Platform','Worked By','Worklog Date','Time Spent (hours)','Worklog Comment']]
    return data_frame

def write_to_excel_file(excel_file_path, data_frame):
    excel_writer = pd.ExcelWriter(excel_file_path, engine='xlsxwriter')
    data_frame.to_excel(excel_writer, index=False)
    excel_writer.save()
    print(excel_file_path + ' has been created.')  
    
if __name__ == '__main__':
    data_frame = read_csv_file(r'C:\Users\User\Documents\Python Scripts\old_sprint.csv')
    write_to_excel_file(r'C:\Users\User\Documents\Python Scripts\new_sprint.xlsx', data_frame)
    

    