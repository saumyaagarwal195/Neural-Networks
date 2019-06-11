#this code extract content from g2crowd of a particular software and populates the Ratings into Excel file

import openpyxl
import tweepy
import nltk
from vaderSentiment.vaderSentiment import SentimentIntensityAnalyzer
import statistics
import math
import numpy as np
import urllib.request, urllib.parse, urllib.error
from urllib.request import urlopen
from bs4 import BeautifulSoup
import ssl
import requests
import json
import re
import sys
import urllib.request
from inscriptis import get_text

like_best=[]
heading=[]
dislikes=[]
recommend=[]
problems=[]
#url = "https://www.g2.com/products/123formbuilder/reviews"
#html = urllib.request.urlopen(url).read().decode('utf-8')

def writing_to_worbook(avg,r,c):
    wbkName = 'D:\\2ndsem\\HIWIJOBS\\first\\test data.xlsx'        #The file should be created before running the code.
    wbk = openpyxl.load_workbook(wbkName)
    wks = wbk['Sheet1']
    someValue = avg
    wks.cell(row=r, column=c).value = someValue
    wbk.save(wbkName)
    wbk.close

def tokenize(t):
    "List all the word tokens (consecutive letters) in a text. Normalize to lowercase."
    return re.findall('[a-z]+', t.lower())

#extracts content from different review into different files.
def extraction(file2):
    file3=open('like.txt','w')
    file4 = open('dislike.txt', 'w')
    file5=open('recommend.txt','w')
    file6 = open('problems.txt', 'w')
    copy3=False
    copy4 = False
    copy5 = False
    copy6 = False
    start0='content'
    start1='What do you like best?'
    end1='What do you dislike?'
    end2='Recommendations to others considering the product'
    end3='What business problems are you solving with the product? What benefits have you realized?'
    end4='* 0 comments'
    #likes=re.search('What do you like best(.+?)What do you dislike',text)

    #likes.txt
    for line in file2:
        if line.strip() == start1:
            copy3 = True
            continue
        elif line.strip() == end1:
            copy3 = False
            continue
        elif copy3:
            file3.write(line)

    #dislikes.txt
    file2.seek(0,0)
    for line in file2:
        if line.strip() == end1:
            copy4 = True
            continue
        elif line.strip() == end2:
            copy4 = False
            continue
        elif copy4:
            file4.write(line)

    file2.seek(0, 0)
    # recommends.txt
    for line in file2:
        if line.strip() == end2:
            copy5 = True
            continue
        elif line.strip() == end3:
            copy5 = False
            continue
        elif copy5:
            file5.write(line)

    # problems.txt
    file2.seek(0, 0)
    for line in file2:
        if line.strip() == end3:
            copy6 = True
            continue
        elif line.strip() == end4:
            copy6 = False
            continue
        elif copy6:
            file6.write(line)

def sentiment_analyzer_scores(sentence):
    score = analyser.polarity_scores(sentence)
    print("{:-<40} {}".format(sentence, str(score)))

def connection():
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE

    user_agent = 'Mozilla/5.0 (Windows; U; Windows NT 5.1; en-US; rv:1.9.0.7) Gecko/2009021910 Firefox/3.0.7'
    #url=input("Enter Software Url - ")
    #headers={'User-Agent':user_agent,}
    hdr = {'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.11 (KHTML, like Gecko) Chrome/23.0.1271.64 Safari/537.11',
       'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
       'Accept-Charset': 'ISO-8859-1,utf-8;q=0.7,*;q=0.3',
       'Accept-Encoding': 'none',
       'Accept-Language': 'en-US,en;q=0.8',
       'Connection': 'keep-alive'}
    html=urllib.request.Request("https://www.g2.com/products/123formbuilder/reviews",None,hdr) #The assembled request
    #resp = urlopen(req)
    resp=urlopen(html).read().decode('utf-8')

    #reads the content of the page like reviews and ratings removing html, javascript content
    text = get_text(resp)

    #write to file everything received in "text"
    file1=open("file1.txt","w")
    file1.write(text)
    file1.close()
    file2=open('file1.txt','r')
    print("file written")
    ##print(type(text))
    ##print(text)

    #extract ratings from the page which is from "Ratings" to "Company Size"
    ratings=(text.split('Ratings'))[1].split('Company Size')[0]
    #print("ratings",type(ratings))
    #print("ratings", ratings)
    rate=[int(s)for s in ratings.split() if s.isdigit()]
    #print(rate)
    count=0
    sum=0
    for i in range(0,len(rate),2):
       sum+=(rate[i]*rate[i+1])
       count+=rate[i+1]
    #finds the average of Ratings
    avg=math.ceil(sum/count)
    print(avg)
    extraction(file2)

    #writes to workbook in excel row 2 column 25 which has heading Ratings
    writing_to_worbook(avg, 2, 25)


analyser = SentimentIntensityAnalyzer("Once an issue has been found, there is no way to tell when it will be resolved")
print(analyser)


connection()
