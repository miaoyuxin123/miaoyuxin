#%%
import  traceback
from ReportSave import ReportSave
import os
import pandas as pd

#%% 基础信息设置

import time
def wechart_file():
    wechart = 'C:\\Users\\25156\\Documents\\WXWork\\1688849964869686\\Cache\\File\\2021-08\\'
    wechart_floder = os.listdir(wechart)
    
    time_file_ser =pd.Series(dtype = 'str')
    for file_name in wechart_floder:
        if '客户基金分析汇总' in file_name:
            file_time = time.localtime(os.stat(wechart + file_name).st_ctime)
            file_time = time.strftime('%Y-%m-%d %H:%M:%S',file_time)
            time_file_ser[file_time] = file_name
    time_file_ser.sort_index(inplace = True)
    print(time_file_ser)
    final_file = wechart + time_file_ser[-1]
    
    customer_df = pd.read_excel(final_file)
    convers_dict = {}
    for col in customer_df.columns:
        convers_dict[col] = str
    customer_df = pd.read_excel(final_file, converters = convers_dict)
    customer_df.set_index('用户ID', inplace = True)
    
    huizong_floder = 'E:\\Project\\2.客户基金分析\\DataSave\\文件汇总\\'
    old_customer_list = os.listdir(huizong_floder)
    new_customer_dict = {}
    for customer_id in customer_df.index:
        if customer_id + '.docx' not in old_customer_list:
            new_customer_list = list(customer_df.loc[customer_id,'基金代码1':].dropna())
            new_customer_dict[customer_id] =  new_customer_list
            
    return new_customer_dict



#%%
base_date = '2021-07-31'
customer_id = '124'
code_list = ['000051']
#code_list = ['','','','','','','','',]
if __name__ == '__main__':            
    if not code_list:
        new_customer_dict = wechart_file()
        print('='*20 + '新增客户' + '='*20)
        print(new_customer_dict)
        print('='*20 + '新增客户' + '='*20)
        for customer_id, code_list in new_customer_dict.items():
            try:
                rs = ReportSave(customer_id, code_list, base_date)
            except Exception as e:
                print(e)
                traceback.print_exc()
                
    else:
        rs = ReportSave(customer_id, code_list, base_date)
            






    

