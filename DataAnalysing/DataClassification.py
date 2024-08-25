import pandas as pd
import numpy as np
from scipy.stats import  chi2_contingency
import openpyxl
from sklearn.preprocessing import StandardScaler
import matplotlib.pyplot as plt

data = pd.read_excel('multireview_metadata_latest_HYX.xlsx')
df = pd.DataFrame(data)
df = df[['Title_EN', 'Avg_rating_EN_gr', 'Num_ratings_EN_gr', 'Rating_az', 'Num_rating_az',
         'db_overall_rating', 'total_raters']]

data_cleaned = df.dropna()
data_cleaned['Norm_Rating_gr'] = (data_cleaned['Avg_rating_EN_gr']-1)/4
data_cleaned['Norm_Rating_az'] = (data_cleaned['Rating_az']-1)/4
data_cleaned['Norm_Rating_db'] = (data_cleaned['db_overall_rating']-1)/9

data_cleaned['Norm_Rating_nums_gr'] = data_cleaned['Num_ratings_EN_gr']/data_cleaned['Num_ratings_EN_gr'].max()
data_cleaned['Norm_Rating_nums_az'] = data_cleaned['Num_rating_az']/data_cleaned['Num_rating_az'].max()
data_cleaned['Norm_Rating_nums_db'] = data_cleaned['total_raters']/data_cleaned['total_raters'].max()

classified_level = []
classified_level.append({'N':data_cleaned['Norm_Rating_nums_gr'].quantile(0.33)*0.33+data_cleaned['Norm_Rating_nums_az'].quantile(0.33)*0.33+data_cleaned['Norm_Rating_nums_db'].quantile(0.33)*0.33,
                         'L':data_cleaned['Norm_Rating_nums_gr'].quantile(0.66)*0.66+data_cleaned['Norm_Rating_nums_az'].quantile(0.66)*0.66+data_cleaned['Norm_Rating_nums_db'].quantile(0.66)*0.66
                         })
data_cleaned['Classify_Score'] = data_cleaned['Norm_Rating_nums_gr']*data_cleaned['Norm_Rating_gr']+data_cleaned['Norm_Rating_nums_az']*data_cleaned['Norm_Rating_az']+data_cleaned['Norm_Rating_nums_db']*data_cleaned['Norm_Rating_db']

data_cleaned['Classify_Type'] = data_cleaned['Classify_Score'].apply(lambda x: (
    'Loved' if x > classified_level[0]['L'] else
    'Normal' if classified_level[0]['N'] <= x <= classified_level[0]['L'] else
    'Dislike'
))


data_cleaned.to_excel('Classification_Type.xlsx',index=False,engine='openpyxl')