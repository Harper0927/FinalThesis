import numpy as py
import pandas as pd
import openpyxl
import matplotlib.pyplot as plt
from scipy.stats import chi2_contingency

data_classified = pd.read_excel('Classification_Type.xlsx',engine='openpyxl')
data_origin = pd.read_excel('multireview_metadata_latest_HYX.xlsx',engine='openpyxl')

#data cleaning
data_origin = data_origin.replace('\n',',',regex = True)

dcf = pd.DataFrame(data_classified)
dof = pd.DataFrame(data_origin)

dof_all_classified = dof[dof['Title_EN'].isin(dcf['Title_EN'])]
tags_o = dof_all_classified['Shelves_gr'].str.replace('[','',regex = True)
tags_o = tags_o.str.replace(']','',regex = True)
tags_f = []
for tags in tags_o:
    tags_list_f=[]
    if '\'' in tags:
        tags_list = tags.split('\'')
        for tag in tags_list:
            temp = tag.split("\\n")[0]
            temp = temp.strip()
            if temp != ',' and temp != '':
                if '>' in temp:
                    temp = temp.split('>')[1].strip()
                if ' ' in temp:
                    temps = temp.split(' ')
                    for item in temps:
                        if item not in tags_list_f and temp != '':
                            tags_list_f.append(item)
                    continue
                if temp not in tags_list_f and temp != '':
                    tags_list_f.append(temp)
    else:
        tags_list = tags.split(',')
        for tag in tags_list:
            if tag != ',' and tag != '':
                if ' ' in tag:
                    tag = tag.split(' ')
                    for item in tag:
                        if item not in tags_list_f and item != '':
                            tags_list_f.append(item)
                    continue
                if tag not in tags_list_f and tag != '':
                    tags_list_f.append(tag)
    tags_f.append(tags_list_f)

dcf['Tags'] = tags_f

# 描述性统计分析

tag_totalnum = 0
tag_totalnum_l = 0
tag_totalnum_n = 0
tag_totalnum_d = 0

tag_dict = dict()
tag_dict_l = dict()
tag_dict_n = dict()
tag_dict_d = dict()

for index, item in dcf.iterrows():
    for tag in item['Tags']:
        tag_totalnum += 1
        if item['Classify_Type'] == 'Loved':
            tag_totalnum_l += 1
            if tag in tag_dict:
                tag_dict[tag] += 1
            else:
                tag_dict[tag] = 1
            if tag in tag_dict_l:
                tag_dict_l[tag] += 1
            else:
                tag_dict_l[tag] = 1
        elif item['Classify_Type'] == 'Normal':
            tag_totalnum_n += 1
            if tag in tag_dict:
                tag_dict[tag] += 1
            else:
                tag_dict[tag] = 1
            if tag in tag_dict_n:
                tag_dict_n[tag] += 1
            else:
                tag_dict_n[tag] = 1
        else:
            tag_totalnum_d += 1
            if tag in tag_dict:
                tag_dict[tag] += 1
            else:
                tag_dict[tag] = 1
            if tag in tag_dict_d:
                tag_dict_d[tag] += 1
            else:
                tag_dict_d[tag] = 1

tag_freq = dict()
tag_freq_loved = dict()
tag_freq_norm = dict()
tag_freq_dislike = dict()

tkeys = list(tag_dict.keys())
tvalues = list(tag_dict.values())
df_tag_dict = pd.DataFrame({'Tag': tkeys, 'Value': tvalues})
df_tag_dict.to_excel('tag_dict.xlsx', index=False, engine='openpyxl')

for key, value in tag_dict.items():
    tag_freq[key] = (value / len(dcf)) * 100
for key, value in tag_dict_l.items():
    tag_freq_loved[key] = (value / len(dcf[dcf['Classify_Type'] == 'Loved'])) * 100
for key, value in tag_dict_n.items():
    tag_freq_norm[key] = (value / len(dcf[dcf['Classify_Type'] == 'Normal'])) * 100
for key, value in tag_dict_d.items():
    tag_freq_dislike[key] = (value / len(dcf[dcf['Classify_Type'] == 'Dislike'])) * 100

sorted_items = sorted(tag_freq.items(), key=lambda item: item[1], reverse=True)
sorted_items_l = sorted(tag_freq_loved.items(), key=lambda item: item[1], reverse=True)
sorted_items_n = sorted(tag_freq_norm.items(), key=lambda item: item[1], reverse=True)
sorted_items_d = sorted(tag_freq_dislike.items(), key=lambda item: item[1], reverse=True)

bar_cate, bar_val = zip(*sorted_items)
bar_cate_l, bar_val_l = zip(*sorted_items_l)
bar_cate_n, bar_val_n = zip(*sorted_items_n)
bar_cate_d, bar_val_d = zip(*sorted_items_d)

plt.figure(figsize=(15, 4))
bars_l = plt.bar(bar_cate_l, bar_val_l, width=0.5)
# 为每个柱子添加数值标签
for bar in bars_l:
    yval = bar.get_height()
    text = format(yval, '.1f')
    plt.text(bar.get_x() + bar.get_width() / 2, yval, text, ha='center', va='bottom', fontsize=4)
plt.title(f'Frequency of Tags in Books Loved')
plt.xlabel('Tags')
plt.ylabel('Freq')
plt.xticks(rotation=90)
plt.margins(x=0.001)
plt.tight_layout()
plt.xticks(fontsize=6)
plt.savefig('sorted_bar_loved_tag_freq.png', dpi=300, bbox_inches='tight', transparent=True)
plt.show()

#plt.figure(figsize=(4, 12))
#bars_l = plt.barh(bar_cate_l, bar_val_l,height=0.5)
#plt.gca().invert_yaxis()
# 为每个柱子添加数值标签
#for bar in bars_l:
#    xval = bar.get_width()
#    text = format(xval, '.1f')
#    plt.text(xval+5, bar.get_y()+bar.get_height()*1.5, text, ha='center', va='bottom', fontsize=6)
#plt.title(f'Frequency of Tags in Books Loved')
#plt.xlabel('Freq')
#plt.ylabel('Tags')
#plt.margins(y=0.001)
#plt.tight_layout()
#plt.yticks(fontsize=8)
#plt.xlim(0,105)
#plt.savefig('sorted_bar_loved_tag_freq3.png', dpi=300, bbox_inches='tight', transparent=True)
#plt.show()

plt.figure(figsize=(15, 4))
bars_n = plt.bar(bar_cate_n, bar_val_n, width=0.5)
# 为每个柱子添加数值标签
for bar in bars_n:
    yval = bar.get_height()
    text = format(yval, '.1f')
    plt.text(bar.get_x() + bar.get_width() / 2, yval, text, ha='center', va='bottom', fontsize=4)
plt.title(f'Frequency of Tags in Books Neutral')
plt.xlabel('Tags')
plt.ylabel('Freq')
plt.xticks(rotation=90)
plt.margins(x=0.001)
plt.tight_layout()
plt.xticks(fontsize=6)
plt.savefig('sorted_bar_neut_tag_freq.png', dpi=300, bbox_inches='tight', transparent=True)
plt.show()

plt.figure(figsize=(15, 4))
bars_d = plt.bar(bar_cate_d, bar_val_d, width=0.5)
# 为每个柱子添加数值标签
for bar in bars_d:
    yval = bar.get_height()
    text = format(yval, '.1f')
    plt.text(bar.get_x() + bar.get_width() / 2, yval, text, ha='center', va='bottom', fontsize=4)
plt.title(f'Frequency of Tags in Books Disliked')
plt.xlabel('Tags')
plt.ylabel('Freq')
plt.xticks(rotation=90)
plt.margins(x=0.001)
plt.tight_layout()
plt.xticks(fontsize=6)
plt.savefig('sorted_bar_disliked_tag_freq.png', dpi=300, bbox_inches='tight', transparent=True)
plt.show()

plt.figure(figsize=(15, 4))
bars = plt.bar(bar_cate, bar_val, width=0.5)
# 为每个柱子添加数值标签
for bar in bars:
    yval = bar.get_height()
    text = format(yval, '.1f')
    plt.text(bar.get_x() + bar.get_width() / 2, yval, text, ha='center', va='bottom', fontsize=4)
plt.title(f'Frequency of Tags in Books')
plt.xlabel('Tags')
plt.ylabel('Freq')
plt.xticks(rotation=90)
plt.margins(x=0.001)
plt.tight_layout()
plt.xticks(fontsize=6)
plt.savefig('sorted_bar_tag_freq.png', dpi=300, bbox_inches='tight', transparent=True)
plt.show()

tag_Remarkable = pd.DataFrame(
    columns=['Tag', 'Frequency', 'Compare to Disliked', 'Compare to Normal', 'Compare to Total', 'Unique'])

for key, value in tag_freq_loved.items():
    tag_Remarkable.loc[len(tag_Remarkable)] = {'Tag': key, 'Frequency': value, 'Compare to Disliked': value,
                                               'Compare to Normal': value,
                                               'Compare to Total': value - tag_freq[key], 'Unique': 0}
    uniq = 1
    if key in tag_freq_dislike:
        tag_Remarkable.loc[len(tag_Remarkable) - 1, 'Compare to Disliked'] = value - tag_freq_dislike[key]
        uniq = 0
    if key in tag_freq_norm:
        tag_Remarkable.loc[len(tag_Remarkable) - 1, 'Compare to Normal'] = value - tag_freq_norm[key]
        uniq = 0
    tag_Remarkable.loc[len(tag_Remarkable) - 1, 'Unique'] = uniq
tag_Remarkable.to_excel('tag_Remarkable.xlsx', index=False, engine='openpyxl')

tag_common_list = []

for key,value in tag_dict.items():
    if key in tag_dict_l:
        if key in tag_dict_n:
            if key in tag_dict_d:
                tag_common_list.append(key)

# OneHot Encoding
OH_tags = dcf['Tags'].explode().str.get_dummies().groupby(level=0).sum()
common_tags = OH_tags.loc[:, tag_common_list]
commen_tags_df = pd.DataFrame(dcf[['Title_EN', 'Classify_Type']])
commen_tags_df = commen_tags_df.join(common_tags)
commen_tags_df.to_excel('Commen_Tag_DataFrame.xlsx', index=False, engine='openpyxl')


