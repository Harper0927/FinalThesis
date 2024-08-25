import time

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from collections import Counter
import openpyxl
# Used for Load model directly
from transformers import AutoTokenizer, AutoModelForSequenceClassification
from transformers import AutoConfig , TFAutoModelForSequenceClassification
from scipy.special import softmax

# Language Detection
from langdetect import detect

#百度通用翻译API,不包含词典、tts语音合成等资源，如有相关需求请联系translate_api@baidu.com
# coding=utf-8
import http.client
import hashlib
import urllib
import random
import json

#用于翻译非英文语句
def TranslationByBaiduAPI(query):
    appid = '20240821002129436'  # 填写你的appid
    secretKey = 'TK1CWRz012IeQjQ6H7uw'  # 填写你的密钥
    httpClient = None
    myurl = '/api/trans/vip/translate'
    fromLang = 'auto'   #原文语种
    toLang = 'en'   #译文语种
    salt = random.randint(32768, 65536)
    q= query
    sign = appid + q + str(salt) + secretKey
    sign = hashlib.md5(sign.encode()).hexdigest()
    myurl = myurl + '?appid=' + appid + '&q=' + urllib.parse.quote(q) + '&from=' + fromLang + '&to=' + toLang + '&salt=' + str(
    salt) + '&sign=' + sign
    try:
        httpClient = http.client.HTTPConnection('api.fanyi.baidu.com')
        httpClient.request('GET', myurl)
        # response是HTTPResponse对象
        response = httpClient.getresponse()
        result_all = response.read().decode("utf-8")
        result = json.loads(result_all)
        transResult = result['trans_result'][0]['dst']
    except Exception as e:
        print (e)
        if httpClient:
            httpClient.close()
        return e
    finally:
        if httpClient:
            httpClient.close()
        return transResult

# Preprocess text (username and link placeholders)
def preprocess(text):
    new_text = []
    for t in text.split(" "):
        t = '@user' if t.startswith('@') and len(t) > 1 else t
        t = 'http' if t.startswith('http') else t
        new_text.append(t)
    return " ".join(new_text)

model_fpath = f"D:/HuggingFaceModel/cardiffnlp/twitter-roberta-base-sentiment-latest"
tokenizer = AutoTokenizer.from_pretrained(model_fpath)
config = AutoConfig.from_pretrained(model_fpath)
model = AutoModelForSequenceClassification.from_pretrained(model_fpath)

comment_list = pd.read_csv("Book_The Three-Body Problem.csv")

df_clist = pd.DataFrame(comment_list)
df_score = pd.DataFrame(columns=['Translation','Sentiment','NegativeScore','NeutralScore','PositiveScore'])
score_array = ['Negative', 'Neutral', 'Positive']
total_process_row = len(df_clist)
for index,row in df_clist.iterrows():
    paragraphs = row['comment_content'].split("\n")
    paragraphs = [item.strip() for item in paragraphs if item.strip()]
    paragraph_array = [0,0,0]
    translation_para = ""
    IsTranslated = False
    totol_paragraphs_count = len(paragraphs)
    for para in paragraphs:
        if para.strip():
            pre_para = preprocess(para)
            try:
                if detect(pre_para) != 'en':
                    pre_para=TranslationByBaiduAPI(pre_para)
                    time.sleep(0.15)
                    print("translated")
            except:
                totol_paragraphs_count -= 1
                continue
            translation_para="{} {}".format(translation_para, pre_para)
            # 截断并填充输入序列
            inputs = tokenizer(pre_para, return_tensors="pt", padding=True, truncation=True, max_length=512)
            output = model(**inputs)
            scores = output[0][0].detach().numpy()
            scores = softmax(scores)
            ranking = np.argsort(scores)
            ranking = ranking[::-1]
            #Negative=0 Neutral=1 Positive=2
            paragraph_array[ranking[0]] += scores[ranking[0]]
            paragraph_array[ranking[1]] += scores[ranking[1]]
            paragraph_array[ranking[2]] += scores[ranking[2]]
    try:
        df_score.loc[len(df_score)] = {'Translation':translation_para,
                                       'Sentiment':score_array[np.argmax(paragraph_array)],
                                       'NegativeScore':paragraph_array[0]/totol_paragraphs_count,
                                       'NeutralScore':paragraph_array[1]/totol_paragraphs_count,
                                       'PositiveScore':paragraph_array[2]/totol_paragraphs_count}
    except:
        df_score.loc[len(df_score)] = {'Translation':translation_para,
                                       'Sentiment': 'Exception',
                                       'NegativeScore': 0,
                                       'NeutralScore': 0,
                                       'PositiveScore': 0}
    print((index+1)," of ",total_process_row)
df_clist = pd.concat([df_clist, df_score], axis=1)
df_clist.to_excel('Goodreads_Comments_with_Sentiment.xlsx',index=False,engine='openpyxl')

