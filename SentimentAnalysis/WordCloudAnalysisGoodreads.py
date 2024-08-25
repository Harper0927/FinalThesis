import pandas as pd
import matplotlib.pyplot as plt
from collections import Counter
from wordcloud import WordCloud
import re
import openpyxl

data = pd.read_excel("Goodreads_Comments_with_Sentiment.xlsx")
df = pd.DataFrame(data)
df_clean = df.loc[df['Sentiment']!='Exception']
df_Posi = df_clean.loc[df_clean['Sentiment']=='Positive']
df_Nega = df_clean.loc[df_clean['Sentiment']=='Negative']
df_Neut = df_clean.loc[df_clean['Sentiment']=='Neutral']

# 计算情感类型的数量
sentiment_counts = df_clean['Sentiment'].value_counts()

# 绘制情感占比的饼状图
#colors = ['#66b3ff', '#ff9999',]
plt.figure(figsize=(8, 8))
wedges, texts, autotexts = plt.pie(sentiment_counts, labels=sentiment_counts.index, autopct='%1.1f%%', startangle=90, textprops=dict(color="w"))

# 添加图例
plt.legend(wedges, sentiment_counts.index, title="Sentiment Type", loc="upper right", bbox_to_anchor=(1,1),fontsize=10)

# 美化饼状图
plt.setp(autotexts, size=20, weight="bold")
plt.title('Proportion of Sentiments in Goodreads Reviews',fontsize = 20)
plt.axis('equal')  # 保证饼状图是圆的
plt.savefig('sentiment_piechart_gr.png',dpi=300, transparent=True)
plt.show()


stopwords_path = '停用词-英文.txt'
# 加载停用词表
with open(stopwords_path, 'r', encoding='utf-8') as f:
    stopwords = set(f.read().strip().split('\n'))
    stopwords.update(['book', 'read'])  # 添加需要去掉的词

# 统计词频
text_total = ' '.join(df_clean['Translation'])
words_total = re.findall(r'\b[A-Za-z]+(?:-[A-Za-z]+)*\b', text_total.lower())
text_pos = ' '.join(df_Posi['Translation'])
words_pos = re.findall(r'\b[A-Za-z]+(?:-[A-Za-z]+)*\b', text_pos.lower())
text_neg = ' '.join(df_Nega['Translation'])
words_neg = re.findall(r'\b[A-Za-z]+(?:-[A-Za-z]+)*\b', text_neg.lower())
text_neu = ' '.join(df_Neut['Translation'])
words_neu = re.findall(r'\b[A-Za-z]+(?:-[A-Za-z]+)*\b', text_neu.lower())
# 过滤停用词
filtered_words_total = [word for word in words_total if word not in stopwords]
filtered_words_pos = [word for word in words_pos if word not in stopwords]
filtered_words_neg = [word for word in words_neg if word not in stopwords]
filtered_words_neu = [word for word in words_neu if word not in stopwords]
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
word_freq_file_path = 'word_frequency_goodreads_'  # 替换为实际文件路径
word_freq_df_pos.to_excel(word_freq_file_path+'Pos.xlsx', index=False)
word_freq_df_neg.to_excel(word_freq_file_path+'Neg.xlsx', index=False)
word_freq_df_neu.to_excel(word_freq_file_path+'Neu.xlsx', index=False)
word_freq_df_total.to_excel(word_freq_file_path+'Total.xlsx', index=False)

# 绘制词云图
wordcloud_total = WordCloud(width=800, height=400, background_color='white').generate_from_frequencies(word_freq_total)
plt.figure(figsize=(10, 5))
plt.imshow(wordcloud_total, interpolation='bilinear')
plt.axis('off')  # 不显示坐标轴
plt.title('Wordcloud of Goodreads Reviews-Total')
plt.savefig('Wordcloud_gr_total.png',dpi=300, transparent=True)
plt.show()

wordcloud_pos = WordCloud(width=800, height=400, background_color='white').generate_from_frequencies(word_freq_pos)
plt.figure(figsize=(10, 5))
plt.imshow(wordcloud_pos, interpolation='bilinear')
plt.axis('off')  # 不显示坐标轴
plt.title('Wordcloud of Goodreads Reviews-Positive')
plt.savefig('Wordcloud_gr_pos.png',dpi=300, transparent=True)
plt.show()

wordcloud_neg = WordCloud(width=800, height=400, background_color='white').generate_from_frequencies(word_freq_neg)
plt.figure(figsize=(10, 5))
plt.imshow(wordcloud_neg, interpolation='bilinear')
plt.axis('off')  # 不显示坐标轴
plt.title('Wordcloud of Goodreads Reviews-Negative')
plt.savefig('Wordcloud_gr_neg.png',dpi=300, transparent=True)
plt.show()

wordcloud_neu = WordCloud(width=800, height=400, background_color='white').generate_from_frequencies(word_freq_neu)
plt.figure(figsize=(10, 5))
plt.imshow(wordcloud_neu, interpolation='bilinear')
plt.axis('off')  # 不显示坐标轴
plt.title('Wordcloud of Goodreads Reviews-Neutral')
plt.savefig('Wordcloud_gr_neu.png',dpi=300, transparent=True)
plt.show()
