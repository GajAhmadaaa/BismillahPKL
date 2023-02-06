import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from wordcloud import WordCloud
import networkx as nx
from Sastrawi.StopWordRemover.StopWordRemoverFactory import StopWordRemoverFactory
import os

cwd = os.getcwd()
img_savepath = f"{cwd}\static\img"

def main(filename):
  global excel, data, question
  excel = pd.ExcelFile(filename)
  data = pd.read_excel(excel, sheet_name=None)

  #Worksheet name extraction
  list_sheet = {}
  i = 1
  for sheet_name in excel.sheet_names:
      list_sheet[str(i)] = sheet_name
      i += 1

  #Extracting question
  question1 = [list(data[i]) for i in excel.sheet_names]
  question = [question1[i][0] for i in range(len(question1))]

  return sorted(list_sheet.items(), key=lambda item: item[0])

def sheet_num(sheet_numb):
  global sheet_number
  sheet_number = sheet_numb

  return sheet_number, excel.sheet_names[sheet_number - 1], question[sheet_number - 1]

def stopword():
  global stopwords_result
  text = list(data[excel.sheet_names[sheet_number - 1]][question[sheet_number - 1]])
  #lowercase and removing punctuation
  new_string = [text[i].translate(str.maketrans(' ', ' ', '1234567890!"#$%&\'()*+,-./:;<=>?@[\\]^_`{|}~\n')).lower() for i in range(len(text))]

  # create stop words removal
  factory = StopWordRemoverFactory()
  stopwords = factory.create_stop_word_remover()
  stopwords_result = []

  # stop words removal process
  for item in new_string:
      output = stopwords.remove(item)
      stopwords_result.append(output)

def wordcloud():
  word_cloud = WordCloud(max_words = 50, width = 1920, height = 1080, background_color = 'white').generate(' '.join(stopwords_result))
  plt.imshow(word_cloud, interpolation='bilinear')
  plt.axis("off")
  plt.savefig(img_savepath + "\wordcloud.jpg")

def freq():
  global most_occur
  from collections import Counter
  Counter = Counter(' '.join(stopwords_result).split())
  most_occur = Counter.most_common(10)

  list_occur = {}
  for i in range(len(most_occur)):
    list_occur[most_occur[i][0]] = most_occur[i][1]

  most_occur = [most_occur[i][0] for i in range(len(most_occur))]

  return sorted(list_occur.items(), reverse=True, key=lambda item: item[1])

def avg():
  avg = []
  for i in stopwords_result:
    avg.append(len(i.split()))
  
  avg_words_per_sentence = round(sum(avg)/len(avg), 2)
  total_word = sum(avg)

  return avg_words_per_sentence, total_word

def ngrams(ngram=2):
  global n_grams
  n_grams = []
  for i in range(len(stopwords_result)):
    converted = [i for i in stopwords_result[i].split()]
    words = converted
    temp=zip(*[words[i:] for i in range(0,ngram)])
    ans=[' '.join(ngram) for ngram in temp]
    n_grams.append(ans)

def network(top=3):
  aarr=[]
  for i in range(len(n_grams)):
    for j in n_grams[i]:
      word = j.split()
      for w in word:
        if w in most_occur[:top]:
          aarr.append(' '.join(word))
        else:
          continue

  aarr = list(dict.fromkeys(aarr))

  arr1=[]
  arr2=[]

  for i in range(len(aarr)):
    p = aarr[i].split()
    arr1.append(p[0])
    arr2.append(p[1])

  # Build a dataframe with your connections
  df = pd.DataFrame({ 'from':arr1, 'to':arr2})
  
  # Build your graph
  G=nx.from_pandas_edgelist(df, 'from', 'to')

  plt.figure(figsize=(16,8))

  # Chart with Custom edges:
  nx.draw_networkx(G, with_labels=True, node_size=1000, node_color="skyblue", node_shape="h", alpha=0.6, linewidths=5, font_size=8, 
          font_color="black", font_weight="normal", width=1, edge_color="grey", pos = nx.spring_layout(G, k=0.6, iterations=100))
  plt.savefig(img_savepath + "\\network.jpg")