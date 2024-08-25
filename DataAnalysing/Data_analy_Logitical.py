import pandas as pd
import numpy as np
from sklearn.preprocessing import LabelEncoder
from sklearn.utils import compute_class_weight
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import classification_report
from scipy import stats
import statsmodels.api as sm
import openpyxl

data = pd.read_excel('Commen_Tag_DataFrame.xlsx',engine='openpyxl')

df = pd.DataFrame(data)

X = df.select_dtypes(include=[np.number])

label_encoder = LabelEncoder()
df['Classify_Type_Num'] = label_encoder.fit_transform(df['Classify_Type'])


y = df['Classify_Type_Num']

# 计算类权重以处理样本不平衡
class_weights = compute_class_weight('balanced', classes=[0,1,2], y=y)
class_weight_dict = {i: class_weights[i] for i in range(len(class_weights))}

# 进行加权的多类别逻辑回归分析
model = LogisticRegression(multi_class='multinomial',solver='newton-cg',tol=0.05,max_iter=2000,class_weight=class_weight_dict)
model.fit(X, y)

y_pred = model.predict(X)
classify_model = classification_report(y, y_pred)
fp = open("Logisitic_Model_Classify_Info.txt","w")
fp.write(classify_model)
fp.close()

# 计算每个样本的预测概率 P(y_i | X_i, \hat{\beta})
predicted_probs = model.predict_proba(X)

# 初始化 Fisher 信息矩阵
fisher_info_matrix = np.zeros((model.coef_.shape[1], model.coef_.shape[1]))
df_p = pd.DataFrame(columns=['Loved','Normal'])
for id_class in model.classes_:
    # 计算 Fisher 信息矩阵
    for i in range(X.shape[0]):
        X_i = X.iloc[i, :].values.reshape(-1, 1)
        P_i = predicted_probs[i, id_class]  # 计算对应类别的 Fisher 信息
        fisher_info_matrix += X_i @ X_i.T * P_i * (1 - P_i)
    # 计算 Fisher 信息矩阵的逆
    cov_matrix = np.linalg.pinv(fisher_info_matrix)#在存在自变量高度共线性（高度相关）or数据样本量不足or权重导致数值不稳定时可能需要用伪逆(pinv)代替(inv)
    # 计算标准误差
    standard_errors = np.sqrt(np.diag(cov_matrix))
    # 计算 z 值
    z_values = model.coef_[id_class] / standard_errors
    # 计算 p 值
    p_values = 2 * (1 - stats.norm.cdf(np.abs(z_values)))
    if id_class == 1:
        df_p['Loved'] = p_values.tolist()
    elif id_class == 2:
        df_p['Normal'] = p_values.tolist()

# 提取分析结果并存储在新的DataFrame中
results_df = pd.DataFrame({
    'Tag': X.columns,
    # 提取系数
    'Coefficient_Loved': model.coef_[1],
    'Coefficient_Normal': model.coef_[2],
})
results_df = pd.concat([results_df, df_p], axis=1)
results_df.to_excel('Logitical_Tags.xlsx',index=False,engine='openpyxl')