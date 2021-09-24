#引入模块
from SmartAnalysis import ReportSave
from CustomerReport import SellBuy
from HandleMergeDf import handle_merge_df
import pandas as pd
import os
import traceback
import numpy as np

#%% 读取基础信息
if __name__ == '__main__':
    #查看本次计算数据
    path_base = 'DataSave\\Customer\\'
    path_save = 'DataSave\\CustomerReport\\'
    path_base_file_list = os.listdir(path_base)

    new_file_list = ['146持仓.xlsx']#'80持仓.xlsx'
    if not new_file_list:
        for file_name in path_base_file_list:
            if file_name[:-5] + '调整计算.xlsx' not in os.listdir(path_save):
                new_file_list.append(file_name)
        print('='*10 + '新增客户' +'='*10)
        print(new_file_list)
        print('='*10 + '新增客户' +'='*10)
         
    #计算并更改表头
    for file_name in new_file_list:
        print('='*10 + file_name[:-5] + '-计算开始' + '='*10)
        try:
            file = path_base + file_name
            cus_df = pd.read_excel(file,converters={'基金代码':str})
            cus_df.set_index('基金代码',inplace = True)
            sb = SellBuy(cus_df)
            merge_df = sb.merge_df.copy()
            merge_df.rename(columns={'current_percent':'原始持仓',
                                     'current_percent_effect':'原始持仓_保留',
                                     'percent_limit':'单只比例限定',
                                     'percent_limit_revised':'单只比例限定_修正',
                                     'final_percent':'最终比例',
                                     'final_asset':'最终资产'}, inplace = True)
            
        except Exception as e:
            print('\r\n\n\n','='*10 + '报错' +'='*10)
            print('客户报错：{}'.format(file_name))
            print(e)
            print('='*10 + file_name[:-5] + '-结束' + '='*10)
            traceback.print_exc(limit = None)
             
        #进一步处理结果
        merge_df_new = handle_merge_df(merge_df)
        part_columns = ['基金名称','基金市值','原始持仓_标记','投资类型(二级分类)',\
                        '新大类','新小类','基金评分','保留原因',\
                        '最终比例','最终资产','买卖金额_取整','操作建议','如何调整']
        merge_df_part = merge_df_new.loc[:,part_columns].copy()
        big_type_df = sb.big_type_df
        if abs(sb.asset_total - merge_df['最终资产'].sum()) < 0.01:
            big_type_df['最终资产'] = sb.asset_total * big_type_df['最终比例']
        else:
            print('总资产计算有误')
        

        
        
        
            
        #%% 生成基金评价
        limit_1 = merge_df['原始持仓_标记']==1
        limit_2 = merge_df['新大类']=='股类'
        
        limit_3 = pd.Series(data = False,index = merge_df.index)
        score_df = pd.read_excel(r'DataBase\基金分类与评分.xlsx',index_col = 0)
        for code in limit_3.index:
            if code in score_df.index:
                #print(code)
                limit_3[code] = True
                
        analysis_df = merge_df_part.loc[limit_1 & limit_2 & limit_3,\
                                      ['基金名称','基金评分','操作建议']]
        analysis_code_list = analysis_df.index.to_list()
              
        cus_id = file_name[:-7]
        base_date = '2021-07-31'
        if not analysis_df.empty:
            rs =  ReportSave(cus_id,analysis_code_list,base_date) 
            analysis_df['基金评价'] = [rs.final_text_dict[code + '.OF'] for code in analysis_df.index]
                    
                    
        #%% 生成基金推荐理由
        limit_6 =merge_df_part['原始持仓_标记']==0
        limit_7 = merge_df_part['新大类']=='另类'
        recomend_df = merge_df_part.loc[limit_6 & (limit_2 | limit_7),['基金名称','基金评分']]
        recommend_code_list = recomend_df.index.to_list()
        
        file = r'E:\Project\4.智能报告\DataBase\推荐理由模板.xlsx'   
        example_df = pd.read_excel(file,converters={'代码':str})
        example_df.drop_duplicates('代码', inplace = True)
        example_df.set_index('代码', inplace = True)
        reason_list = []
        for code in recommend_code_list:
            if code in example_df.index:
                reason_list.append(example_df.loc[code,'基金评价'])
            else:
                reason_list.append(np.nan)
        
        recomend_df['综合建议'] = '强烈建议购买' 
        recomend_df['分析'] = reason_list 
        #recomend_df.reindex()
#%%数据保存
        excel_name = path_save + file_name[:-5] +'调整计算.xlsx'
        with pd.ExcelWriter(excel_name) as writer:
            merge_df_part.to_excel(writer, sheet_name = '持仓调整_简版')
            big_type_df.reindex(['股类','债类','另类']).to_excel(writer, sheet_name = '大类汇总')
            analysis_df.to_excel(writer, sheet_name ='现有产品分析')
            recomend_df.to_excel(writer, sheet_name ='推荐理由')
            
            merge_df_new.to_excel(writer, sheet_name = '持仓调整_详细')
            sb.small_type_df.reindex(['债类','主动','指数','qdii','另类']).to_excel(writer, sheet_name = '小类汇总')
            print('='*10 + file_name[:-5] + '-计算结束' + '='*10)
          
            
            
            
            
            
            
            
            

