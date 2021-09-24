#%% 引入模块
from docx import Document
import datetime
from pandas.tseries.offsets import DateOffset
import pandas as pd
from pandas import read_excel
import numpy as np
from WindPy import*
w.start()


#%% 提取数据
class ReportKey():
    def __init__(self,code,base_date,index_code,score_dict):
        self.code = code
        self.base_date = base_date
        self.index_code = index_code
        self.score_dict = score_dict
        self.inint_func()
        
    def inint_func(self):
        self.init_date()
        self.score()
        self.info()
        self.rate()
        self.rank()
        self.index()
        self.std_maxdown()
        self.manager()
        self.company()
        self.pepb()
        
    def init_date(self):

        base_date_dt = datetime.strptime(self.base_date,'%Y-%m-%d')
        self.base_date_dt = base_date_dt
        self.year_start_dt = (base_date_dt + DateOffset(month = 1,day = 1)).date()
        #self.month_offset_3 = (base_date_dt + DateOffset(months = -3)).date()#注意是months，复数
        #self.month_offset_6 = (base_date_dt + DateOffset(months = -6)).date()
        self.month_offset_12 = (base_date_dt + DateOffset(months = -12)).date()
        self.month_offset_36 = (base_date_dt + DateOffset(months = -36)).date()
        self.setup_date_dt = w.wss(self.code, "fund_setupdate").Data[0][0].date()
        
        self.year_start = datetime.strftime(self.year_start_dt,'%Y-%m-%d')#今年
        #self.date3 = datetime.strftime(self.month_offset_3,'%Y-%m-%d')#3
        #self.date6 = datetime.strftime(self.month_offset_6,'%Y-%m-%d')#6
        self.date12 = datetime.strftime(self.month_offset_12,'%Y-%m-%d')#12
        self.date36 = datetime.strftime(self.month_offset_36,'%Y-%m-%d')#36
        self.setup_date = datetime.strftime(self.setup_date_dt,'%Y-%m-%d')#成立以来

    def score(self):
        for key,value in self.score_dict.items():
            if self.code in value.index:
                self.fund_score = value.loc[self.code,'总评分（满100分）']

    
    
    def info(self):
        """基础信息"""
        info_dict = {}
        data = w.wss(self.code, "fund_investobject,fund_investconception,fund_riskreturn_characters,fund_firstinvesttype,fund_investtype","tradeDate=20210720")
        info_dict['投资目标'] = data.Data[0]
        info_dict['投资理念'] = data.Data[1]
        info_dict['风险收益特征'] = data.Data[2]
        info_dict['投资类型一级'] = data.Data[3]
        info_dict['投资类型二级'] = data.Data[4]
        self.info_dict = info_dict
    
    def rate(self):
        """收益率"""
        rate_dict = {}
        data = w.wss(self.code, "return_ytd,return_3m,return_6m,return_1y,return_3y,return_std","annualized=0;tradeDate=20210720")
        rate_dict['今年以来回报'] =data.Data[0]
# =============================================================================
#         rate_dict['近3月回报'] =data.Data[1]
#         rate_dict['近6月回报'] =data.Data[2]
# =============================================================================
        rate_dict['近1年回报'] =data.Data[3]
        rate_dict['近3年回报'] =data.Data[4]
        rate_dict['成立以来回报'] =data.Data[5]
        self.rate_dict = rate_dict


    def rank(self):
        """排名信息"""
        self.year_start_rank = w.wss(self.code, "peer_fund_return_rank_per,peer_fund_return_rank_prop_per",\
                           "startDate={};endDate={};fundType=3".format(self.year_start, self.base_date)).Data
# =============================================================================
#         self.date3_rank = w.wss(self.code, "peer_fund_return_rank_per,peer_fund_return_rank_prop_per",\
#                            "startDate={};endDate={};fundType=3".format(self.date3, self.base_date)).Data
#         self.date6_rank = w.wss(self.code, "peer_fund_return_rank_per,peer_fund_return_rank_prop_per",\
#                            "startDate={};endDate={};fundType=3".format(self.date6, self.base_date)).Data
# =============================================================================
        self.date12_rank = w.wss(self.code, "peer_fund_return_rank_per,peer_fund_return_rank_prop_per",\
                           "startDate={};endDate={};fundType=3".format(self.date12, self.base_date)).Data
        self.date36_rank = w.wss(self.code, "peer_fund_return_rank_per,peer_fund_return_rank_prop_per",\
                           "startDate={};endDate={};fundType=3".format(self.date36, self.base_date)).Data
        self.setup_date_rank = w.wss(self.code, "peer_fund_return_rank_per,peer_fund_return_rank_prop_per",\
                           "startDate={};endDate={};fundType=3".format(self.setup_date, self.base_date)).Data
        rank_dict = {}
        rank_dict['今年以来排名'] = self.year_start_rank
# =============================================================================
#         rank_dict['近3月回报排名'] = self.date3_rank
#         rank_dict['近6月回报排名'] = self.date6_rank
# =============================================================================
        rank_dict['近1年回报排名'] = self.date12_rank
        rank_dict['近3年回报排名'] = self.date36_rank
        rank_dict['成立以来回报排名'] = self.setup_date_rank
        self.rank_dict = rank_dict

    def index(self):
        """对比的指数信息"""
        self.index_code_name = w.wss(self.index_code, "sec_name").Data[0]
        self.base_date_nav = w.wss(self.index_code, "fund_setupdate,close","tradeDate={};priceAdj=U;cycle=D".format(self.base_date)).Data[1][0]
        self.year_start_nav = w.wss(self.index_code, "fund_setupdate,close","tradeDate={};priceAdj=U;cycle=D".format(self.year_start)).Data[1][0]
# =============================================================================
#         self.date3_nav = w.wss(self.index_code, "fund_setupdate,close","tradeDate={};priceAdj=U;cycle=D".format(self.date3)).Data[1][0]
#         self.date6_nav = w.wss(self.index_code, "fund_setupdate,close","tradeDate={};priceAdj=U;cycle=D".format(self.date6)).Data[1][0]
# =============================================================================
        self.date12_nav = w.wss(self.index_code, "fund_setupdate,close","tradeDate={};priceAdj=U;cycle=D".format(self.date12)).Data[1][0]
        self.date36_nav = w.wss(self.index_code, "fund_setupdate,close","tradeDate={};priceAdj=U;cycle=D".format(self.date36)).Data[1][0]
        self.setup_date_nav = w.wss(self.index_code, "fund_setupdate,close","tradeDate={};priceAdj=U;cycle=D".format(self.setup_date)).Data[1][0]
        
        index_dict = {}
        index_dict['指数今年以来回报'] = self.base_date_nav / self.year_start_nav - 1
# =============================================================================
#         index_dict['指数近3月回报'] = self.base_date_nav /self.date3_nav - 1
#         index_dict['指数近6月回报'] = self.base_date_nav /self.date6_nav - 1
# =============================================================================
        try:
            index_dict['指数近1年回报'] = self.base_date_nav /self.date12_nav - 1
        except:
            index_dict['指数近1年回报'] = np.nan
        try:
            index_dict['指数近3年回报'] = self.base_date_nav /self.date36_nav - 1
        except:
            index_dict['指数近3年回报'] = np.nan
        index_dict['指数成立以来回报'] = self.base_date_nav / self.setup_date_nav - 1
        self.index_dict = index_dict
    


    def std_maxdown(self):
        """波动率和最大回撤"""
        self.year_start_std_maxdown = w.wss(self.code, "risk_stdevyearly,risk_maxdownside",\
                                  "startDate={};endDate={};period=2;returnType=1".format(self.year_start, self.base_date)).Data
# =============================================================================
#         self.date3_std_maxdown = w.wss(self.code, "risk_stdevyearly,risk_maxdownside",\
#                                   "startDate={};endDate={};period=2;returnType=1".format(self.date3, self.base_date)).Data  
#         self.date6_std_maxdown = w.wss(self.code, "risk_stdevyearly,risk_maxdownside",\
#                                   "startDate={};endDate={};period=2;returnType=1".format(self.date6, self.base_date)).Data     
# =============================================================================
        self.date12_std_maxdown = w.wss(self.code, "risk_stdevyearly,risk_maxdownside",\
                                  "startDate={};endDate={};period=2;returnType=1".format(self.date12, self.base_date)).Data     
        self.date36_std_maxdown = w.wss(self.code, "risk_stdevyearly,risk_maxdownside",\
                                  "startDate={};endDate={};period=2;returnType=1".format(self.date36, self.base_date)).Data     
        self.setup_date_std_maxdown = w.wss(self.code, "risk_stdevyearly,risk_maxdownside",\
                                  "startDate={};endDate={};period=2;returnType=1".format(self.setup_date, self.base_date)).Data     
        
        std_maxdown_dict = {}
        std_maxdown_dict['今年以来波动回撤'] = self.year_start_std_maxdown
# =============================================================================
#         std_maxdown_dict['近3月回报波动回撤'] = self.date3_std_maxdown
#         std_maxdown_dict['近6月回报波动回撤'] = self.date6_std_maxdown
# =============================================================================
        std_maxdown_dict['近1年回报波动回撤'] = self.date12_std_maxdown
        std_maxdown_dict['近3年回报波动回撤'] = self.date36_std_maxdown
        std_maxdown_dict['成立以来回报波动回撤'] = self.setup_date_std_maxdown 
        self.std_maxdown_dict = std_maxdown_dict
    

    def manager(self):
        """基金经理信息"""
        manager_data = w.wss(self.code, "fund_fundmanageroftradedate,fund_manager_totalnetasset,fund_manager_fundno",\
                                  "tradeDate={};unit=1;order=1".format(self.base_date))
        manager_data2 = w.wss(self.code, "fund_manager_startdate,fund_manager_managerworkingyears","order=1")
        
        manager_dict = {}
        manager_dict['基金经理'] = manager_data.Data[0][0]
        manager_dict['管理规模'] = manager_data.Data[1][0]/100000000
        manager_dict['基金数量'] = manager_data.Data[2][0]
        manager_dict['任职日期'] = manager_data2.Data[0][0].date()
        manager_dict['任职年限'] = manager_data2.Data[1][0]
        self.manager_dict =manager_dict

    def company(self):
        """基金公司信息"""
        company_df = pd.read_excel(r'D:\BackTesting\DataManager\基金公司数据（7.22）.xlsx', index_col = 0)
        company_name = w.wss(self.code, "fund_mgrcomp").Data[0][0]
        company_dict = {}
        company_dict['基金公司'] = company_name
        company_dict['管理总规模'] = company_df.loc[company_name,'总规模']
        company_dict['成立日期'] = datetime.strptime(company_df.loc[company_name,'成立日期'],'%Y-%m-%d').date()
        company_dict['排名'] = '{}/{}'.format(company_df.index.get_loc(company_name) + 1,len(company_df))
        company_dict['排名百分比'] = (company_df.index.get_loc(company_name) + 1) / len(company_df)
        self.company_dict = company_dict


    def pepb(self):
        """获取指数pepb估值"""
        def guzhi(num):
            if num < 0.1:
                return '极度低估'
            elif num < 0.2:
                return '低估'
            elif num < 0.4:
                return '正常偏低'
            elif num < 0.6:
                return '正常'
            elif num < 0.8:
                return '正常偏高'
            elif num < 0.9:
                return '高估'            
            else:
                return '极度高估'
            
        if self.info_dict['投资类型二级'][0] == '被动指数型基金':
            pepb_dict = {}
            pepb_path = 'E:\\Project\\1.基金评分\\DataBase\\PEPB基础数据.xlsx'
            pe_df = pd.read_excel(pepb_path, sheet_name = 'PE',index_col = 0)
            pb_df = pd.read_excel(pepb_path, sheet_name = 'PB',index_col = 0)
            pe_percent_df = pd.read_excel(pepb_path, sheet_name = 'PE百分比',index_col = 0)
            pb_percent_df = pd.read_excel(pepb_path, sheet_name = 'PB百分比',index_col = 0)

            
            pepb_dict['追踪指数代码'] = self.index_code
            if self.index_code in pe_percent_df.index:
                pepb_dict['追踪指数名称'] = pe_percent_df.loc[self.index_code,'指数名称']
                
                pepb_dict['追踪指数pe倍数'] = pe_df.loc[self.index_code,:][-1]
                pepb_dict['追踪指数pb倍数'] = pb_df.loc[self.index_code,:][-1]
                
                pepb_dict['追踪指数pe倍数_历史平均'] = pe_df.loc[self.index_code,:][1:].mean()
                pepb_dict['追踪指数pb倍数_历史平均'] = pb_df.loc[self.index_code,:][1:].mean()
                
                pepb_dict['追踪指数pe百分位'] = pe_percent_df.loc[self.index_code,:][-1]
                pepb_dict['追踪指数pb百分位'] = pb_percent_df.loc[self.index_code,:][-1]
                
                
                pepb_dict['pe估值状态'] = guzhi(pepb_dict['追踪指数pe百分位'])
                pepb_dict['pb估值状态'] = guzhi(pepb_dict['追踪指数pb百分位'])
            
                self.pepb_dict = pepb_dict
            else:
                self.pepb_dict = pepb_dict = {}
                
        else:
            self.pepb_dict = pepb_dict = {}
#%%
# =============================================================================
# score_dict = {}
# path = 'C:\\Users\\25156\\Desktop\\组合分析\\7.31\\'
# for name in ['股票+混合+增强','另类','债券','指数','qdii']:
#     score_dict[name] = pd.read_excel(path +name + '.xlsx',index_col = 0)  
# code_list = '000051.OF'
# customer_id = '998'
# index_code = '000300.SH'
# base_date = '2021-07-31'
# code_type = '被动指数型基金'
# 
# rk = ReportKey(code_list, base_date,index_code,score_dict)
# =============================================================================


#%%
class ReportText():
    def __init__(self, rk):
        self.rk = rk
        self.init_func()
        
    def init_func(self):
        self.fund_info()
        self.fund_nav()
        self.fund_risk()
        self.fund_manager()
        self.fund_company()
        self.fund_pepb()
        self.final_limit()
    
    def text_format(self,text,*args):
        try:
            return text.format(*args)
        except TypeError:
            return ''
        except Exception as e:
            print(e)
            
    def fund_info(self):
        text_format = self.text_format
        setup_date_dt = self.rk.setup_date_dt
        info_dict = self.rk.info_dict
        
        try:
            self.t0 = '评分：{:.2f}\r\n'.format(self.rk.fund_score)
        except:
            self.t0 = '评分：--\r\n'
        self.t1 = text_format('本基金是{}，',info_dict['投资类型一级'][0])
        self.t2 = text_format('属于{}。',info_dict['投资类型二级'][0])
        self.t3 = text_format('基金成立{}年{}月{}日，',setup_date_dt.year,setup_date_dt.month,setup_date_dt.day)
        self.t4 = text_format('投资目标是{}',info_dict['投资目标'][0])#投资目标
        self.t5 = text_format('投资理念是{}',info_dict['投资理念'][0])#投资理念
        #self.fund_info_list = [t0 , t1 ,t2 ,t3 ,t4 , t5]
        
    
    def fund_nav(self):
        text_format = self.text_format
        base_date_dt = self.rk.base_date_dt
        rate_dict = self.rk.rate_dict
        rank_dict = self.rk.rank_dict
        index_code_name = self.rk.index_code_name
        index_dict = self.rk.index_dict
        
        self.t6 = '\r\n历史业绩：本基金与相关指数{}对比来看，'.format(index_code_name[0])
        t7_1 = text_format('2021年初至{}年{}月{}日，',base_date_dt.year,base_date_dt.month,base_date_dt.day)
        t7_2 = text_format('同期本基金上涨{:.2f}%，',rate_dict['今年以来回报'][0])
        t7_3 = text_format('相关指数{}上涨{:.2f}%，',index_code_name[0],index_dict['指数今年以来回报']*100)
        self.t7 = t7_1 +  t7_3 + t7_2
        self.t8 = text_format('最近3年相关指数{}上涨{:.2f}%，本基金上涨{:.2f}%，',\
                              index_code_name[0],index_dict['指数近3年回报']*100,rate_dict['近3年回报'][0], )
        self.t9 = text_format('基金自{}成立以来，相关指数{}收益{:.2f}%，本基金累计收益{:.2f}%，',\
                              self.t3[4:-1],index_code_name[0],index_dict['指数成立以来回报']*100, rate_dict['成立以来回报'][0],)  
        self.t10 = '从基金的同类排名上来看，'
        self.t11 = text_format('在同类排名为{}，处于行业前{:.2f}%；',rank_dict['今年以来排名'][0][0],rank_dict['今年以来排名'][1][0])
        self.t12 = text_format('在同类排名为{}，处于行业前{:.2f}%；',rank_dict['近1年回报排名'][0][0],rank_dict['近1年回报排名'][1][0])
        self.t13 = text_format('在同类排名为{}，处于行业前{:.2f}%；',rank_dict['近3年回报排名'][0][0],rank_dict['近3年回报排名'][1][0])
        self.t14 = text_format('在同类排名为{}，处于行业前{:.2f}%。',rank_dict['成立以来回报排名'][0][0],rank_dict['成立以来回报排名'][1][0])
        #self.fund_nav_list = [t6 , self.t7 ,self.t8 ,t9 ,t10 , self.t11 ,self.t12 ,self.t13 ,t14]
     
    
    def fund_risk(self):
        text_format = self.text_format
        std_maxdown_dict = self.rk.std_maxdown_dict
        
        self.t15_1 = '\r\n风险分析：从基金的波动率和最大回撤来看，'
        self.t15_2 = text_format('基金最近1年的净值波动率为{:.2f}%，最大回撤为{:.2f}%；',std_maxdown_dict['近1年回报波动回撤'][0][0], std_maxdown_dict['近1年回报波动回撤'][1][0])
        self.t16 = text_format('最近3年的净值波动率为{:.2f}%，最大回撤为{:.2f}%；',std_maxdown_dict['近3年回报波动回撤'][0][0], std_maxdown_dict['近3年回报波动回撤'][1][0])
        self.t17 =text_format('成立以来的净值波动率为{:.2f}%，最大回撤为{:.2f}%。',std_maxdown_dict['成立以来回报波动回撤'][0][0], std_maxdown_dict['成立以来回报波动回撤'][1][0])
            
        #self.fund_risk_list = [t15 ,self.t15_2 ,self.t16 ,t17]

    
    def fund_manager(self):
        manager_dict = self.rk.manager_dict
        
        self.t18 = '\r\n基金经理分析：现任基金经理为{}，于{}年{}月{}日开始管理本基金，总计管理{:.2f}年。'.\
        format(manager_dict['基金经理'],manager_dict['任职日期'].year,manager_dict['任职日期'].month,manager_dict['任职日期'].day,manager_dict['任职年限'])
        self.t19 = '该基金经理目前管理规模{:.2f}亿元，'.format(manager_dict['管理规模'])
        self.t20 = '管理基金{}只。'.format(manager_dict['基金数量'])
        
        #过滤1：基金经理过滤
        if self.rk.info_dict['投资类型二级'][0] == '被动指数型基金':
            self.t18 = self.t19 = self.t20 = ''


    def fund_company(self):
        company_dict = self.rk.company_dict
        self.t21 = '\r\n基金公司分析：{}成立于{}年，目前管理总规模{:.2f}亿元，在所有公募基金公司排名为{}，处于前{:.2f}%的位置。'.format(\
        company_dict['基金公司'] ,company_dict['成立日期'].year,company_dict['管理总规模'],\
            company_dict['排名'], company_dict['排名百分比']*100)
        #self.fund_company_list = [t21]
        
    def fund_pepb(self):
        text_format = self.text_format
        pepb_dict = self.rk.pepb_dict
        if (self.rk.info_dict['投资类型二级'][0] == '被动指数型基金') & (bool(pepb_dict)):
            t22_1 = text_format('\r\n估值分析：本指数基金跟踪的指数名称{}，',pepb_dict['追踪指数名称'])
            t22_2 = text_format('代码{}。',pepb_dict['追踪指数代码'])
            self.t22 = t22_1 + t22_2
            self.t23 = text_format('从跟踪指数历史pe（市盈率）来看，pe的历史平均估值为{:.2f}倍，当前估值{:.2f}倍，位于{:.2f}%的位置，属于{}状态；',\
                               pepb_dict['追踪指数pe倍数_历史平均'],pepb_dict['追踪指数pe倍数'],pepb_dict['追踪指数pe百分位']*100,pepb_dict['pe估值状态'])
            self.t24 = text_format('从跟踪指数历史pb（市净率）来看，pb的历史平均估值为{:.2f}倍，当前估值{:.2f}倍，位于{:.2f}%的位置，属于{}状态。',\
                               pepb_dict['追踪指数pb倍数_历史平均'],pepb_dict['追踪指数pb倍数'],pepb_dict['追踪指数pb百分位']*100,pepb_dict['pb估值状态'])
        else:
            self.t22 = self.t23 = self.t24 = ''
  
        
    def final_limit(self):
        """过滤2：日期过滤"""
        rk = self.rk  
        if rk.setup_date_dt > rk.month_offset_36:
            self.t8 = ''
            self.t13 = ''
            self.t16 = ''
            
        if rk.setup_date_dt > rk.month_offset_12:
            self.t12 = ''
            self.t15_2 = ''
            
            
        if rk.setup_date_dt > datetime(2021,1,1).date():
            self.t7 = ''
            self.t11 = ''
            
        final_list = [self.t1,self.t2,self.t3,self.t4,self.t5,
                      self.t6,self.t7,self.t11,
                      self.t8,self.t13,
                      self.t9,self.t14,
                      self.t22,self.t23,self.t24,
                      self.t15_1,self.t15_2,self.t16,self.t17,
                      self.t18,self.t19,self.t20,
                      self.t21] 
            
        final_text = ''
        for t in final_list:
            if ('None' in t) or ('nan' in t):#过滤3：空值过滤
                t = ''   
            final_text = final_text + t
        self.final_text = final_text
 

    
    
    
    
    
    