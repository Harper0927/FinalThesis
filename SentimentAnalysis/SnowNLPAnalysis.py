import pandas as pd
from snownlp import SnowNLP
import matplotlib.pyplot as plt
from collections import Counter
from wordcloud import WordCloud
import re

# 设置中文字体显示
plt.rcParams["font.sans-serif"] = ["SimHei"]  # 用来显示中文
plt.rcParams["axes.unicode_minus"] = False  # 用来显示负号

# 加载数据
file_path = 'douban_short_comment.xlsx'  # 替换为实际文件路径
stopwords_path = '停用词.txt'  # 替换为实际停用词文件路径
df = pd.read_excel(file_path)

# 清洗数据，去除空值，并确保评论内容为字符串类型
df['评论内容'] = df['评论内容'].astype(str).fillna('')

# 对评论内容进行情感分析
df['情感得分'] = df['评论内容'].apply(lambda x: SnowNLP(x).sentiments)
df['情感类型'] = df['情感得分'].apply(lambda x: 'Positive' if x >= 0.66 else ('Neutral'if x >= 0.33 else 'Negative'))
df_Posi = df.loc[df['情感类型']=='Positive']
df_Nega = df.loc[df['情感类型']=='Negative']
df_Neut = df.loc[df['情感类型']=='Neutral']
# 计算情感类型的数量
sentiment_counts = df['情感类型'].value_counts()

# 绘制情感占比的饼状图
#colors = ['#66b3ff', '#ff9999']
plt.figure(figsize=(8, 8))
wedges, texts, autotexts = plt.pie(sentiment_counts, labels=sentiment_counts.index, autopct='%1.1f%%', startangle=140, textprops=dict(color="w"))

# 添加图例
plt.legend(wedges, sentiment_counts.index, title="Sentiment Type", loc="upper right", bbox_to_anchor=(1, 1))

# 美化饼状图
plt.setp(autotexts, size=20, weight="bold")
plt.title('Proportion of Sentiments in Douban Reviews',fontsize = 20)
plt.axis('equal')  # 保证饼状图是圆的
plt.savefig('sentiment_piechart_db.png',dpi=300, transparent=True)
plt.show()

# 保存结果到新文件
output_file_path = 'Douban_Reviews_with_Sentiment.xlsx'  # 替换为实际文件路径
df.to_excel(output_file_path, index=False)
print(f"结果已保存到: {output_file_path}")
