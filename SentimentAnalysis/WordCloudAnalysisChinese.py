import pandas as pd
import jieba
from snownlp import SnowNLP
import matplotlib.pyplot as plt
from collections import Counter
from wordcloud import WordCloud
import re

# 设置中文字体显示
plt.rcParams["font.sans-serif"] = ["SimHei"]  # 用来显示中文
plt.rcParams["axes.unicode_minus"] = False  # 用来显示负号

# 加载数据
file_path = 'Douban_Reviews_with_Sentiment.xlsx'  # 替换为实际文件路径
stopwords_path = '停用词.txt'  # 替换为实际停用词文件路径
df = pd.read_excel(file_path)

# 加载停用词表
with open(stopwords_path, 'r', encoding='utf-8') as f:
    stopwords = set(f.read().strip().split('\n'))
    stopwords.update(['部', '读','太'])  # 添加需要去掉的词

# 清洗数据，去除空值，并确保评论内容为字符串类型
df['评论内容'] = df['评论内容'].astype(str).fillna('')
df_Posi = df.loc[df['情感类型']=='Positive']
df_Nega = df.loc[df['情感类型']=='Negative']
df_Neut = df.loc[df['情感类型']=='Neutral']
# 统计词频
text = ' '.join(df['评论内容'])
words = " ".join(jieba.cut(text, cut_all=False)).split(" ")
words = [item.strip() for item in words if item.strip()]
filtered_words_total = [word for word in words if word not in stopwords]
filtered_words_total = [s for s in filtered_words_total if not re.search(r'[A-Za-z0-9]', s)]

text_pos = ' '.join(df_Posi['评论内容'])
words_posu = " ".join(jieba.cut(text_pos, cut_all=False)).split(" ")
words_pos = [item.strip() for item in words_posu if item.strip()]
filtered_words_pos = [word for word in words_pos if word not in stopwords]
filtered_words_pos = [s for s in filtered_words_pos if not re.search(r'[A-Za-z0-9]', s)]

text_neg = ' '.join(df_Nega['评论内容'])
words_neg = " ".join(jieba.cut(text_neg, cut_all=False)).split(" ")
words_neg = [item.strip() for item in words_neg if item.strip()]
filtered_words_neg = [word for word in words_neg if word not in stopwords]
filtered_words_neg = [s for s in filtered_words_neg if not re.search(r'[A-Za-z0-9]', s)]

text_neu = ' '.join(df['评论内容'])
words_neu = " ".join(jieba.cut(text_neu, cut_all=False)).split(" ")
words_neu = [item.strip() for item in words_neu if item.strip()]
filtered_words_neu = [word for word in words_neu if word not in stopwords]
filtered_words_neu = [s for s in filtered_words_neu if not re.search(r'[A-Za-z0-9]', s)]

word_freq_total = Counter(filtered_words_total)
word_freq_pos = Counter(filtered_words_pos)
word_freq_neg = Counter(filtered_words_neg)
word_freq_neu = Counter(filtered_words_neu)
# 创建词频数据框
word_freq_df_total = pd.DataFrame(word_freq_total.items(), columns=['Word', 'Frequency'])
word_freq_df_total['Occurrence Probability'] = (word_freq_df_total['Frequency'] / word_freq_df_total['Frequency'].sum()).round(6)
word_freq_df_pos = pd.DataFrame(word_freq_pos.items(), columns=['Word', 'Frequency'])
word_freq_df_pos['Occurrence Probability'] = (word_freq_df_pos['Frequency'] / word_freq_df_pos['Frequency'].sum()).round(6)
word_freq_df_neg = pd.DataFrame(word_freq_neg.items(), columns=['Word', 'Frequency'])
word_freq_df_neg['Occurrence Probability'] = (word_freq_df_neg['Frequency'] / word_freq_df_neg['Frequency'].sum()).round(6)
word_freq_df_neu = pd.DataFrame(word_freq_neu.items(), columns=['Word', 'Frequency'])
word_freq_df_neu['Occurrence Probability'] = (word_freq_df_neu['Frequency'] / word_freq_df_neu['Frequency'].sum()).round(6)

# 保存词频数据框
word_freq_file_path = 'word_frequency_douban_'  # 替换为实际文件路径
word_freq_df_pos.to_excel(word_freq_file_path+'Pos.xlsx', index=False)
word_freq_df_neg.to_excel(word_freq_file_path+'Neg.xlsx', index=False)
word_freq_df_neu.to_excel(word_freq_file_path+'Neu.xlsx', index=False)
word_freq_df_total.to_excel(word_freq_file_path+'Total.xlsx', index=False)


# 绘制词云图，确保显示中文
wordcloud = WordCloud(font_path='simhei.ttf', width=800, height=400, background_color='white').generate_from_frequencies(word_freq_total)
plt.figure(figsize=(10, 5))
plt.imshow(wordcloud, interpolation='bilinear')
plt.axis('off')  # 不显示坐标轴
plt.title('Wordcloud of Douban Reviews-Total')
plt.savefig('Wordcloud_db_total.png',dpi=300, transparent=True)
plt.show()

wordcloud_pos = WordCloud(font_path='simhei.ttf',width=800, height=400, background_color='white').generate_from_frequencies(word_freq_pos)
plt.figure(figsize=(10, 5))
plt.imshow(wordcloud_pos, interpolation='bilinear')
plt.axis('off')  # 不显示坐标轴
plt.title('Wordcloud of Douban Reviews-Positive')
plt.savefig('Wordcloud_db_pos.png',dpi=300, transparent=True)
plt.show()

wordcloud_neg = WordCloud(font_path='simhei.ttf',width=800, height=400, background_color='white').generate_from_frequencies(word_freq_neg)
plt.figure(figsize=(10, 5))
plt.imshow(wordcloud_neg, interpolation='bilinear')
plt.axis('off')  # 不显示坐标轴
plt.title('Wordcloud of Douban Reviews-Negative')
plt.savefig('Wordcloud_db_neg.png',dpi=300, transparent=True)
plt.show()

wordcloud_neu = WordCloud(font_path='simhei.ttf',width=800, height=400, background_color='white').generate_from_frequencies(word_freq_neu)
plt.figure(figsize=(10, 5))
plt.imshow(wordcloud_neu, interpolation='bilinear')
plt.axis('off')  # 不显示坐标轴
plt.title('Wordcloud of Douban Reviews-Neutral')
plt.savefig('Wordcloud_db_neu.png',dpi=300, transparent=True)
plt.show()
