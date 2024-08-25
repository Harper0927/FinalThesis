import pandas as pd
from scipy.stats import fisher_exact
import openpyxl
import numpy as np

# 导入数据
data = pd.read_excel('Commen_Tag_DataFrame.xlsx',engine='openpyxl')

df = pd.DataFrame(data)

df_tags = df.select_dtypes(include=[np.number])
df['Classify_Binary'] = df['Classify_Type'].apply(lambda x: 'high' if x == 'Loved' else 'low')
df_LN = df[df['Classify_Type'].isin(['Loved','Normal'])]
df_LD = df[df['Classify_Type'].isin(['Loved','Dislike'])]
df_ND = df[df['Classify_Type'].isin(['Normal','Dislike'])]
dfLN_tags = df_LN.select_dtypes(include=[np.number])
dfLN_tags = dfLN_tags.drop(columns=dfLN_tags.columns[(dfLN_tags==0).all()])
dfLD_tags = df_LD.select_dtypes(include=[np.number])
dfLD_tags = dfLD_tags.drop(columns=dfLD_tags.columns[(dfLD_tags==0).all()])
dfND_tags = df_ND.select_dtypes(include=[np.number])
dfND_tags = dfND_tags.drop(columns=dfND_tags.columns[(dfND_tags==0).all()])

df_fisher_result = pd.DataFrame(columns=['Tag','Odds Ratio-L-N','p-L-N','Odds Ratio-L-D','p-L-D','Odds Ratio-N-D','p-N-D','Odds Ratio','p'])

# Fisher精确检验：对每个标签计算与受欢迎程度的关联
for tag in df_tags.columns:
    # 构建2x2列联表
    contingency_table = pd.crosstab(df[tag], df['Classify_Binary'])
    # 进行Fisher精确检验
    odds_ratio, p_value = fisher_exact(contingency_table)
    df_fisher_result.loc[len(df_fisher_result)] = {'Tag':tag,'Odds Ratio-L-N':np.nan,'p-L-N':np.nan,'Odds Ratio-L-D':np.nan,'p-L-D':np.nan,'Odds Ratio-N-D':np.nan,'p-N-D':np.nan,'Odds Ratio':odds_ratio,'p':p_value}
###
##df_fisher_result['Influence'] = df.assign(Influence='Not Remarkable')['Influence']
#Loved_Ratio_Total = len(df[df['Classify_Type']=='Loved'])/len(df)
#Remakeable_result = df_fisher_result[df_fisher_result['p']<=0.05]

#for tag in Remakeable_result['Tag']:
#    tempdf = df[df[tag]==1]
#    if len(tempdf[tempdf['Classify_Type']=='Loved'])/len(tempdf) > Loved_Ratio_Total:
#        (df_fisher_result[df_fisher_result['Tag'] == tag])['Influence'] = 'Positive'
#    else:
#        (df_fisher_result[df_fisher_result['Tag'] == tag])['Influence']  = 'Negative'
###
for tag in dfLN_tags.columns:
    contingency_table_LN = pd.crosstab(df_LN[tag], df_LN['Classify_Type'])
    odds_ratio_LN, p_value_LN = fisher_exact(contingency_table_LN)
    df_fisher_result.loc[df_fisher_result['Tag'] == tag ,['Odds Ratio-L-N','p-L-N']] = [odds_ratio_LN,p_value_LN]
for tag in dfLD_tags.columns:
    contingency_table_LD = pd.crosstab(df_LD[tag], df_LD['Classify_Type'])
    odds_ratio_LD, p_value_LD = fisher_exact(contingency_table_LD)
    df_fisher_result.loc[df_fisher_result['Tag'] == tag ,['Odds Ratio-L-D','p-L-D']] = [odds_ratio_LD,p_value_LD]
for tag in dfND_tags.columns:
    contingency_table_ND = pd.crosstab(df_ND[tag], df_ND['Classify_Type'])
    odds_ratio_ND, p_value_ND = fisher_exact(contingency_table_ND)
    df_fisher_result.loc[df_fisher_result['Tag'] == tag,['Odds Ratio-N-D','p-N-D']] = [odds_ratio_ND,p_value_ND]
df_fisher_result.to_excel('Fisher_Tags.xlsx',index=False,engine='openpyxl')
