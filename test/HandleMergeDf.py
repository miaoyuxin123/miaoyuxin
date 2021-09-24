
import pandas as pd
import numpy as np

#%%
def handle_merge_df(merge_df):
    merge_df = merge_df.copy()
    interval_dict = {'股类':{'month':6,'interval':12},
                     '债类':{'month':3,'interval':6},
                     '另类':{'month':4,'interval':8},}

    position_df = merge_df[merge_df['原始持仓_标记'] == 1].copy()
    por_df = merge_df[merge_df['原始持仓_标记'] == 0].copy()
    
    position_df['买卖金额'] = (position_df['最终资产'] - position_df['基金市值'].fillna(0)).apply(lambda x:x if x<0 else 0)
    position_df['买卖金额_取整'] = position_df['买卖金额'].apply(lambda x:int(x))
    por_df['买卖金额'] = (por_df['最终资产'] -por_df['基金市值'].fillna(0)).apply(lambda x:x if x>0 else 0)
    por_df['买卖金额_取整'] = por_df['买卖金额'].apply(lambda x:int(x))
    
    
    
    #持仓卖出
    position_df['操作建议'] =(position_df['买卖金额']/position_df['基金市值']).apply\
        (lambda x:'建议保留' if abs(x) < 0.000001 else('建议卖出' if x== -1 else '部分保留'))
    operate_advice_list = []
    for idx in position_df.index:
        advice = position_df.loc[idx,'操作建议']
        big_type = position_df.loc[idx,'新大类']
        if advice == '建议保留':
            operate_advice_list.append('保留原金额')
        elif advice == '建议卖出':
            if big_type in ['股类']:
                operate_advice_list.append('择机一次性卖出或随定投节奏逐步卖出')
            elif big_type in ['债类','另类']:
                operate_advice_list.append('随定投节奏逐步卖出')
            else:
                raise ValueError('大类分类错误，无法生成操作建议')
        elif advice == '部分保留':
            sell_asset = (position_df.loc[idx,'买卖金额_取整'])*-1
            if big_type in ['股类']:
                operate_advice_list.append('择机一次性卖出或随定投节奏逐步卖出{}元'.format(sell_asset))
            elif big_type in ['债类','另类']:
                operate_advice_list.append('随定投节奏逐步卖出{}元'.format(sell_asset))
            else:
                raise ValueError('大类分类错误，无法生成操作建议')
        else:
            print('错误：操作建议有误')
    
    position_df['如何调整'] = operate_advice_list
    
    #组合买入
    por_df['操作建议'] = por_df['买卖金额'].apply(lambda x:'建议新增' if int(x)>0 else np.nan )
    operate_advice_list = []
    for idx in por_df.index:
        advice = por_df.loc[idx,'操作建议']
        big_type = por_df.loc[idx,'新大类']
        month = interval_dict[big_type]['month']
        interval = interval_dict[big_type]['interval']
        money = int(por_df.loc[idx,'买卖金额_取整']/interval)
        
        if advice == '建议新增':
            if big_type in ['股类']:
                operate_advice_list.append('建议{}个月内双周定投{}次，每次金额{}元'.format(month,interval,money))
            elif big_type in ['债类','另类']:
                operate_advice_list.append('建议{}个月内双周定投{}次，每次金额{}元'.format(month,interval,money))
        
        else:
            operate_advice_list.append(np.nan)

    por_df['如何调整'] = operate_advice_list
    merge_df = pd.concat([position_df, por_df],axis = 0)
    return merge_df
    

        
        