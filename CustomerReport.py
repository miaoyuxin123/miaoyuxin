#%% 引入模块
import pandas as pd
import numpy as np

#%% 临时可删除
# =============================================================================
# path_base = 'E:\\Project\\3.客户报告\\Customer\\'
# 
# file = path_base + '115持仓.xlsx'
# cus_df = pd.read_excel(file,converters={'基金代码':str})
# cus_df.set_index('基金代码',inplace = True)
# =============================================================================
#%%
class ReadData():
    def __init__(self, cus_df):
        self.cus_df = cus_df
        self.init_ReadData()
    
    def init_ReadData(self):
        self.columns_add_list = [ '投资类型(二级分类)', '新大类','新小类', 
         '跟踪指数代码', '基金评分', '强制保留', '普通保留', 
         '同组合基金', '指数白名单', '同组合指数', '是否评分50', '保留原因', '保留标记', 
         'current_percent', 'current_percent_effect', 'percent_limit', 'percent_limit_revised',
       'final_percent', 'final_asset']
        
        self.cus_df_func()
        self.por_df_func()
        self.type_and_index()
    
    def cus_df_func(self):
        """读取基础数据"""
        cus_df = self.cus_df
        cus_cash = cus_df.loc['现金总计','基金名称']
        cus_risk = cus_df.loc['风险级别','基金名称']
        cus_df = cus_df.iloc[:-8,:]
        cus_df_bool = cus_df.isnull()
        bool_final = cus_df_bool.iloc[:,0]
        for col in cus_df_bool.columns:#去除所有列为空
            bool_final = bool_final & cus_df_bool[col]
        cus_df = cus_df.loc[~bool_final, :].copy()  
        cus_df.loc[:,self.columns_add_list] = np.nan
        self.cus_cash = cus_cash
        self.cus_risk = cus_risk
        self.cus_df = cus_df
    
    def por_df_func(self):
        por_df =  pd.read_excel(r'E:\Project\3.客户报告\DataBase\基金组合.xlsx',converters={'代码':str})
        por_df.rename(columns = {'代码':'基金代码','名称':'基金名称'}, inplace = True)
        por_df.dropna(inplace = True)
        por_df = por_df[por_df['风险级别'] == self.cus_risk].copy()
        por_df.set_index('基金代码', inplace = True)
        por_df.loc[:,self.columns_add_list] = np.nan
        self.por_df = por_df        
        
    
    def type_and_index(self):
        cus_df = self.cus_df
        por_df = self.por_df
        type_df = pd.read_excel(r'DataBase\基金分类与评分.xlsx', index_col = 0)
        type_name_list = ['投资类型(二级分类)', '新大类','新小类','跟踪指数代码','基金评分']
        for code in cus_df.index:
            if code in type_df.index:
                cus_df.loc[code, type_name_list] =\
                    type_df.loc[code,type_name_list]
            elif not pd.isnull(cus_df.loc[code,'强制保留类型（一级）']):
                cus_df.loc[code,'新大类'] = cus_df.loc[code,'强制保留类型（一级）']
            elif not pd.isnull(cus_df.loc[code,'普通保留类型（一级）']):
                cus_df.loc[code,'新大类'] = cus_df.loc[code,'普通保留类型（一级）']
                if pd.isnull(cus_df.loc[code,'普通保留类型（二级）']):
                    print('资产的普通保留类型（二级）缺失：{}'.format(code))
                    return
                cus_df.loc[code,'新小类'] = cus_df.loc[code,'普通保留类型（二级）']
            else: 
                raise ValueError('资产类型不可划分：{}'.format(code))
                #print('资产类型不可划分：{}'.format(code))
                #return
        por_df.loc[:,type_name_list] = \
            type_df.loc[por_df.index,type_name_list] 
            
        

       
        
#rd = ReadData(cus_df)  
#cus_df = rd.cus_df
#print(rd.__dict__)   

#%% 添加保留原因
class SaveDrop(ReadData):
    def __init__(self,cus_df):
        self.cus_df = cus_df
        self.init_ReadData()
        self.init_SaveDrop()
        
        
    def init_SaveDrop(self):
        self.must_retained()
        self.same_por_fund()
        self.same_index_fund()
        self.score_50()
        self.save_reason()

        
    def must_retained(self):
        cus_df = self.cus_df       
        cus_df.loc[~pd.isnull(cus_df['强制保留类型（一级）']),'强制保留'] = '是'
        cus_df.loc[pd.isnull(cus_df['强制保留类型（一级）']),'强制保留'] = '否'  
        cus_df.loc[~pd.isnull(cus_df['普通保留类型（一级）']),'普通保留'] = '是'
        cus_df.loc[pd.isnull(cus_df['普通保留类型（一级）']),'普通保留'] = '否'
        
        self.cus_df = cus_df
    
    def same_por_fund(self):
        por_fund_bool = []
        for code in self.cus_df.index:
            if code in self.por_df.index:
                por_fund_bool.append('是')
            else:
                por_fund_bool.append('否')
        self.cus_df['同组合基金'] = por_fund_bool
        
           
    
    def same_index_fund(self):
        fund_index_bool = []
        white_list_bool = []
        white_list = ['000905.SH','000016.SH','000688.SH','399006.SZ','399673.SZ','931643.CSI',]
                      #中证500     上证50      科创50       创业板指    创业板50    科创创业50
        
        for fund_index in self.cus_df['跟踪指数代码']:
            if pd.isnull(fund_index):
                fund_index_bool.append('否')
                white_list_bool.append('否')
                continue
            
            if fund_index in self.por_df['跟踪指数代码'].values:
                fund_index_bool.append('是')
            else:
                fund_index_bool.append('否')
            if fund_index in white_list:
                white_list_bool.append('是')
            else:
                white_list_bool.append('否')
                
        self.cus_df['指数白名单'] = white_list_bool
        self.cus_df['同组合指数'] = fund_index_bool
        
 
    def score_50(self):
        score_50 = self.cus_df['基金评分']>=50
        self.cus_df.loc[score_50,'是否评分50'] = '是'
        self.cus_df.loc[~score_50,'是否评分50'] = '否'

       
    def save_reason(self):
        cus_df = self.cus_df
        limit_must = cus_df['强制保留'] == '是'
        limit_ordinary = cus_df['普通保留'] == '是'
        limit_por = cus_df['同组合基金'] == '是'
        limit_white = cus_df['指数白名单'] == '是'
        limit_index = cus_df['同组合指数'] == '是'
        limit_score = cus_df['是否评分50'] == '是'
        
        cus_df.loc[limit_score,'保留原因'] = '评分保留'
        cus_df.loc[limit_index,'保留原因'] = '同组合指数'
        cus_df.loc[limit_white,'保留原因'] = '指数白名单'
        cus_df.loc[limit_por,'保留原因'] = '同组合基金'
        cus_df.loc[limit_ordinary,'保留原因'] = '普通保留'
        cus_df.loc[limit_must,'保留原因'] = '强制保留'
        
        limit_save = (limit_must |limit_ordinary|limit_por |limit_white|limit_index | limit_score)
        cus_df.loc[limit_save,'保留标记'] = 1
        cus_df.loc[~(limit_save),'保留标记'] = 0
        
  
#sd =    SaveDrop(cus_df)  
#print(sd.__dict__)   

#%% 添加组合基金
class PositionPercent(SaveDrop):
    def __init__(self,cus_df):
        self.cus_df = cus_df       
        self.init_ReadData()
        self.init_SaveDrop()
        self.init_PositionPercent()
        
    def init_PositionPercent(self):
        self.add_portfolio()
        self.position_percent()
        self.small_type()
        self.big_type()
        self.small_type_add()

    def add_portfolio(self):
        por_df = self.por_df
        cus_df = self.cus_df
        
        for code in por_df.index:
            code_index = por_df.loc[code,'跟踪指数代码']
            if code in cus_df.index :
                por_df.loc[code,'保留原因'] = '客户持同基金'
                por_df.loc[code,'保留标记'] = 1
                    
            elif not pd.isnull(code_index):
                if code_index in list(cus_df['跟踪指数代码']) :
                    por_df.loc[code,'保留原因'] = '客户持同指数'
                    por_df.loc[code,'保留标记'] = 1
                else:
                    por_df.loc[code,'保留原因'] = '组合新增'
                    por_df.loc[code,'保留标记'] = 1
            else:
                por_df.loc[code,'保留原因'] = '组合新增'
                por_df.loc[code,'保留标记'] = 1
    
    def position_percent(self):
        cus_df = self.cus_df
        self.asset_total = self.cus_df['基金市值'].sum() + self.cus_cash
        #限定单只持仓百分比
        single_percent_limit = 0.15
        cus_df['current_percent'] = (cus_df['基金市值'] / self.asset_total).fillna(0)
        cus_df['current_percent_effect'] = cus_df['current_percent'] *cus_df['保留标记']
        cus_df['percent_limit'] = single_percent_limit
        cus_df['percent_limit_revised'] = cus_df.loc[:,['current_percent_effect','percent_limit']].min(axis = 1)
        limit_must = cus_df['强制保留'] == '是'
        cus_df.loc[limit_must,'percent_limit_revised'] = cus_df.loc[limit_must,'current_percent']
        
        
    def small_type(self):
        risk_invest_dict = {
             '高风险':{'主动':(5, 0.35),'指数':(4, 0.35),'债类':(2, 0.1),'另类':(1,0.05 ),'qdii':(2,0.15 ),},
           '中高风险':{'主动':(3,0.25),'指数':(3,0.27 ),'债类':(3, 0.28),'另类':(1,0.05 ),'qdii':(2,0.15 ),},
           '中低风险':{'主动':(3,0.2),'指数':(2, 0.2),'债类':(3,0.46 ),'另类':(1,0.04),'qdii':(2,0.1 ),},
             '低风险':{'主动':(2,0.1),'指数':(2, 0.18),'债类':(5,0.7 ),'另类':(1,0.02 ),'qdii':(0, 0),},
            }

        small_type_df = pd.DataFrame(columns = ['新大类','基金数量','标准比例'])
        for key,value in risk_invest_dict[self.cus_risk].items():#[self.cus_risk].items():
            if key in ['主动','指数','qdii']:
                big_type = '股类'
            elif key in ['债类']:
                big_type = '债类'
            elif key in ['另类']:
                big_type = '另类'
            small_type_df.loc[key,:] = [big_type] + list(value)
        self.small_type_df = small_type_df
        
    def big_type(self):
        cus_df = self.cus_df
        small_type_df = self.small_type_df
        #计算强留和非强比例
        big_type_ser = small_type_df['标准比例'].groupby(small_type_df['新大类']).sum()
        big_type_df = pd.DataFrame(big_type_ser)  
        big_type_df['强留比例'] = cus_df['percent_limit_revised'].groupby(cus_df['强制保留类型（一级）']).sum()
        big_type_df.fillna(0, inplace = True)
        big_type_df['非强比例'] = (big_type_df['标准比例'] - big_type_df['强留比例']).apply(lambda x: x if x>=0 else 0)
        #计算给与
        percent_remainder = 1 - big_type_df['强留比例'].sum()
        
        big_type_df.loc['债类','非强比例_实际'] = min(big_type_df.loc['债类','非强比例'], percent_remainder) 
        percent_remainder = percent_remainder - big_type_df.loc['债类','非强比例_实际']
        big_type_df.loc['股类','非强比例_实际'] = min(big_type_df.loc['股类','非强比例'], percent_remainder) 
        percent_remainder = percent_remainder - big_type_df.loc['股类','非强比例_实际']
        big_type_df.loc['另类','非强比例_实际'] = min(big_type_df.loc['另类','非强比例'], percent_remainder)
        #计算最终
        big_type_df['最终比例'] = big_type_df['强留比例'] + big_type_df['非强比例_实际']
        self.big_type_df = big_type_df
        
        if abs(big_type_df['最终比例'].sum() - 1) < 0.01:
            print("大类计算无误")
        else:
            print('大类计算有误')
            return
        
    def small_type_add(self):
        """small_type_df增加 非强比例_实际，用了big_type_df的非强比例_实际"""
        small_type_df = self.small_type_df
        big_type_df = self.big_type_df
        for small_type in small_type_df.index:
            small_data_biaozhun = small_type_df.loc[small_type,'标准比例']
            if small_type in ['主动','指数','qdii']:
                big_data_biaozhun = big_type_df.loc['股类','标准比例']
                big_data_geiyu = big_type_df.loc['股类','非强比例_实际']
                
            elif small_type in ['债类']:
                big_data_biaozhun = big_type_df.loc['债类','标准比例']
                big_data_geiyu = big_type_df.loc['债类','非强比例_实际']
                
            elif small_type in ['另类']:
                big_data_biaozhun = big_type_df.loc['另类','标准比例']
                big_data_geiyu = big_type_df.loc['另类','非强比例_实际']
                
            effect_percent = big_data_geiyu*(small_data_biaozhun/big_data_biaozhun)
            small_type_df.loc[small_type,'非强比例_实际'] = effect_percent        

#pp = PositionPercent(cus_df)   

#%% 资产卖出和买入
class SellBuy(PositionPercent):
    def __init__(self, cus_df):
        self.cus_df = cus_df   
        self.init_ReadData()
        self.init_SaveDrop()
        self.init_PositionPercent()
        self.init_SellBuy()
            
        
    def init_SellBuy(self):
        self.sell()
        self.buy()
        self.merge_sell_buy()
        
    def sell_for_score(self,df):
        """传入一个按照评分顺序保留的df，按照保留评分更高原则计算"""
        
        small_type = df['新小类'][0]
        #print(small_type)
        percent_total_limit = self.small_type_df.loc[small_type,'非强比例_实际']
        df.sort_values('基金评分', ascending = False,inplace = True)
        ser = df['percent_limit_revised']
        
        if not ser.empty:
            if ser.sum() > percent_total_limit:  
                data_list = []
                percent_total = 0
                for data in ser:
                    if percent_total + data <= percent_total_limit:
                        percent_total = percent_total + data
                    else:
                        data = percent_total_limit - percent_total
                        percent_total = percent_total_limit
                    data_list.append(data)
                df['final_percent'] = data_list
            else:
                df['final_percent'] = df['percent_limit_revised']
        return df

        
    
    def sell_for_averge(self,df):
        """传入一个评分带空值的df，按照保留值更平均的原则计算"""
        small_type = df['新小类'][0]
        percent_total_limit = self.small_type_df.loc[small_type,'非强比例_实际']
        df.sort_values('percent_limit_revised',inplace = True)
        ser = df['percent_limit_revised']
        
        df_delta = pd.DataFrame(ser)
        df_delta['num'] = range(len(ser), 0, -1)
        df_delta['delta'] = (ser - ser.shift(1)).fillna(0)
        
        df_final = pd.DataFrame(index = df_delta.index)
        if ser[0]*len(ser) > percent_total_limit:
            final_percent = percent_total_limit / len(ser)
            df_final['base'] = final_percent
            
        else:
            df_final['base'] = ser[0]
            for idx in df_delta.index:
                percent_total = df_final.sum().sum()
                num = df_delta.loc[idx,'num']
                delta = df_delta.loc[idx,'delta']
                if (percent_total + num*delta) <=  percent_total_limit:
                    percent_total += num*delta
                    df_final.loc[idx:,idx] = delta
                    
                else:
                    delta = (percent_total_limit - percent_total)/num
                    df_final.loc[idx:,idx] = delta
        df.loc[:,'final_percent'] = df_final.sum(axis = 1)
        return df
   
            
    
    def sell(self):
        """如果都有评分，按照评分卖；如果含无评分的，按照平均保留来卖"""
        cus_df = self.cus_df
        cus_df_part = cus_df[cus_df['强制保留'] == '否'].copy()
        
        def select_sell_method(df):
            if df.empty:
                return df
            if True in list(pd.isnull(df['基金评分'])):#基金评分有空值
                return self.sell_for_averge(df)
            else:
                return self.sell_for_score(df)
        
        zhudong_df = cus_df_part.loc[cus_df_part['新小类'] == '主动',\
                     ['新小类','基金评分','percent_limit_revised']].copy()
        zhishu_df = cus_df_part.loc[cus_df_part['新小类'] == '指数',\
                    ['新小类','基金评分','percent_limit_revised']].copy()
        zhailei_df = cus_df_part.loc[cus_df_part['新小类'] == '债类',\
                    ['新小类','基金评分','percent_limit_revised']].copy()
        linglei_df = cus_df_part.loc[cus_df_part['新小类'] == '另类',\
                    ['新小类','基金评分','percent_limit_revised']].copy()
        qdii_df = cus_df_part.loc[cus_df_part['新小类'] == 'qdii',\
                    ['新小类','基金评分','percent_limit_revised']].copy()
   
        df_dict = {'主动':zhudong_df,'指数':zhishu_df,
                   '债类':zhailei_df,'另类':linglei_df,'qdii':qdii_df}
        for key,value in df_dict.items():
            new_value = select_sell_method(value)
            if not new_value.empty:
                df_dict[key] = new_value
                cus_df_part.loc[new_value.index,'final_percent'] = new_value.loc[:,'final_percent']
        self.df_dict = df_dict  
        self.cus_df_part = cus_df_part
        cus_df['final_percent'] = cus_df['current_percent']
        cus_df.loc[cus_df_part.index,'final_percent'] = cus_df_part['final_percent']

        
    def buy(self):
        small_type_df = self.small_type_df
        cus_df_part = self.cus_df_part
        por_df = self.por_df
        
        #small_type_df = sb.small_type_df
        #cus_df_part = sb.cus_df_part
        #por_df = sb.por_df
        
        small_type_df['非强持仓保留'] = cus_df_part['final_percent'].groupby(cus_df_part['新小类']).sum()
        small_type_df.fillna(0, inplace = True)
        small_type_df['购买比例'] = small_type_df['非强比例_实际'] - small_type_df['非强持仓保留']
           
        por_df = self.por_df
        for code in por_df.index:#组合的基金代码
            code_index = por_df.loc[code,'跟踪指数代码']
            if not pd.isnull(code_index):#组合基金跟踪指数不为空
                #cus_df_part_part的两种创建情况
                cus_df_part_part = cus_df_part[cus_df_part['跟踪指数代码'] == code_index]
                if code_index == 'NDX.GI':
                    limit_1 = cus_df_part['跟踪指数代码'] == code_index
                    limit_2 = cus_df_part['跟踪指数代码'] == 'XNDX.O'
                    cus_df_part_part = cus_df_part[limit_1 | limit_2]
                #修改组合的名称
                if not cus_df_part_part.empty:#组合存在指数和持仓相同
                    for idx in cus_df_part_part.index:
                        code_position =  idx
                        code_position_name = cus_df_part_part.loc[code_position,'基金名称']
                        if '增强' not in code_position_name:#非增强
                            por_df.rename(index = {code:code_position},inplace = True)
                            por_df.loc[code_position,'基金名称'] = code_position_name
                            por_df.loc[code_position,'跟踪指数代码'] = cus_df_part_part.loc[code_position,'跟踪指数代码']
        
        #判断客户已持仓同基金或者同指数
        por_df['已持仓'] = 0
        for code in por_df.index:
            code_index = por_df.loc[code,'跟踪指数代码']
            if code in cus_df_part.index:
                por_df.loc[code,'已持仓'] = cus_df_part.loc[code,'final_percent'].sum()
            if code_index in list(cus_df_part['跟踪指数代码']):#对应指数不为空
                por_df.loc[code,'已持仓'] = \
                    cus_df_part.loc[cus_df_part['跟踪指数代码'] == code_index,'final_percent'].sum()
        
        def buy_for_average(ser,limit):
            if (ser.sum() >= limit) or (ser.empty):
                return ser
            else:
                ser_min = ser.min()
                ser_second_min = ser[~(ser == ser_min)].min()
                delta_min = ser_second_min - ser_min
                ser_min_num = len(ser[ser == ser_min])
                if ser.sum() + delta_min*ser_min_num <= limit:
                    ser[ser == ser_min] = ser[ser == ser_min] + delta_min
                    return buy_for_average(ser,limit)
                else:
                    delta_min = (limit - ser.sum())/ser_min_num
                    ser[ser == ser_min] = ser[ser == ser_min] + delta_min
                    return ser
           
               
        for small_type in ['主动','指数','另类','qdii','债类']:
           type_limit = por_df['新小类'] == small_type
           por_df_position = por_df.loc[type_limit, ['已持仓']].copy()
           percent_total = por_df_position['已持仓'].sum() + small_type_df.loc[small_type,'购买比例']
           positon_ser = por_df_position['已持仓'].copy()#不带copy会修改por_df_position
           por_df_position['最终持仓'] = buy_for_average(positon_ser,percent_total)
           por_df_position['组合购买'] = por_df_position['最终持仓'] - por_df_position['已持仓']
           por_df.loc[por_df_position.index,'final_percent'] = por_df_position['组合购买']
        
     
    def merge_sell_buy(self):
        #添加原始持仓标记
        por_df_new = pd.DataFrame(self.por_df, columns = self.cus_df.columns)
        insert_local = list(por_df_new.columns).index('普通保留类型（二级）') + 1
        self.cus_df.insert(insert_local,'原始持仓_标记',1)
        por_df_new.insert(insert_local,'原始持仓_标记',0)
       
        #合并表格
        merge_df = pd.concat([self.cus_df, por_df_new], axis = 0)
        merge_df['final_asset'] = merge_df['final_percent']*self.asset_total
        self.merge_df = merge_df
        if abs(merge_df['final_percent'].sum() - 1) < 0.0001:
            print('最终持仓无误')
        else:
            print('错误提醒：最终持仓错误,最终持仓总计{}'.format(merge_df['final_percent'].sum()))
            
#sb = SellBuy(cus_df)

















