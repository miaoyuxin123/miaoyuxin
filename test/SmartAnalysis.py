from ReportKey import ReportKey,ReportText
#from docx import Document
import os
import  traceback
import pandas as pd 
from datetime import date,datetime,timedelta

from WindPy import*
w.start()
#%%
class ReportSave:
    def __init__(self, customer_id, code_list, base_date):
        self.customer_id = customer_id
        self.code_list = code_list
        self.base_date = base_date
        
        self.score_dict()
        self.code_type()
        self.get_document()
        
    def score_dict(self):
        score_dict = {}
        path = 'C:\\Users\\25156\\Desktop\\组合分析\\7.31\\'
        for name in ['股票+混合+增强','另类','债券','指数','qdii']:
            score_dict[name] = pd.read_excel(path +name + '.xlsx',index_col = 0)  
        self.score_dict = score_dict
        
    def code_type(self):
        #code_list = ['002656','588000','000828','001643','000051','000248','260104','000071']
        code_list = self.code_list
        code_wind_list = [code + '.OF' for code in code_list]
        code_type_data = w.wss(code_wind_list, "name_official,fund_firstinvesttype,fund_investtype")
        type_df = pd.DataFrame(code_type_data.Data,\
                               index = ['基金名称','基金类型（一级）','基金类型（二级）'],\
                               columns = code_type_data.Codes).T
        type_index_dict = {
            '股票型基金':'000300.SH',
            '混合型基金':'000300.SH',
            '债券型基金':'H11009.CSI',
            '另类投资基金':'AUFI.WI',
            '国际(QDII)基金':'892400.MI'
            }
        type_df['index_code'] = [type_index_dict[typ] for typ in type_df['基金类型（一级）']]
        beidong_list = type_df[type_df['基金类型（二级）'] == '被动指数型基金'].index.to_list()
        type_df.loc[type_df['基金类型（二级）'] == '被动指数型基金','index_code'] = \
            w.wss(beidong_list, "fund_trackindexcode").Data[0]
        
        print(type_df)
        self.type_df = type_df
     
    
    def get_document(self):
        document_dict = {}
        print('\r\n正在计算：')
        #original_doc = Document()
        for code_wind,index_code in zip(self.type_df.index, self.type_df['index_code']):
            code_name = self.type_df.loc[code_wind,'基金名称']
            head_line = code_name +'-' + code_wind[:6] #标题
            #original_doc.add_heading(head_line, level = 0)
            print(code_wind, index_code)
            try:
                RK = ReportKey(code_wind,self.base_date, index_code, self.score_dict)
                document_dict[code_wind] = RK  #提取关键字
                for key,value in RK.__dict__.items():
                    if key != 'score_dict':
                        original_text = str(key) + '：' +str(value)
                        #original_doc.add_paragraph(original_text)
            except Exception as e:
                print(e)
                traceback.print_exc()
            
        #document = Document() #生成一个word文本 
        self.final_text_dict = {}
        for code_wind,value in document_dict.items():
            code_name = self.type_df.loc[code_wind,'基金名称']
            head_line = code_name +'-' + code_wind[:6] #标题
            #document.add_heading(head_line, level = 0)
            
            try:
                RT = ReportText(value)#关键字合成文本
                self.final_text_dict[code_wind] = RT.final_text 
                #document.add_paragraph(RT.final_text)

            except Exception as e:
                print(e)
                traceback.print_exc()

# =============================================================================
# base_date = '2021-07-31'
# customer_id = '124'
# code_list = ['000051','510300']
# rs = ReportSave(customer_id, code_list, base_date)
# 
# =============================================================================




























