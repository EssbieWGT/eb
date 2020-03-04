from newspaper import Article
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import datetime
import os
import sys

#set directory to current folder
relpath = os.path.realpath(sys.argv[0])
dname = os.path.dirname(relpath)
os.chdir(dname)

#This is another way to do it, but not one that works when freezing code in py2app
# abspath = os.path.abspath(__file__)
# dname = os.path.dirname(abspath)
# os.chdir(dname)

#load excel file into script
# df = pd.read_excel('/Users/essbie/Documents/wcsNewsApp/articles.xlsx', sheet_name='Sheet1')
df = pd.read_excel('articles.xlsx', sheet_name='Sheet1')
print("Loaded Articles")

#sort excel file by category rank and story rank
df.sort_values(by=['categoryOrder','articleOrder'], inplace=True)

#variable for today's date to sub in missing dates
today = datetime.datetime.now().now()

#pull article data
newsPull = []
for i in df["url"]:
    article = Article(i)
    article.download()
    article.parse()
    newsPull.append(article)

#create empty lists for article data
Date = []
Title = []
Authors = []
Text = []

print("Processing Articles")

#pull data out of articles and append lists
for i in newsPull:
    if i.publish_date != None:
        Date.append(i.publish_date)
    else:
        Date.append(today)
    Title.append(i.title)
    if i.authors != []:
        Authors.append(i.authors[0])
    else:
        Authors.append("Not Available")
    Text.append(i.text)

#add article data to df
df.insert(5,"Date", Date)
df.insert(6,"Title", Title)
df.insert(7,"Authors", Authors)
df.insert(8,"Text", Text)

#get unique values from category column
uniqueValues = df.category.unique()

#create read/write plain text file to load data into
file1 = open("./EarlyBird.txt","w+",encoding="UTF-8")

#adding the headlines section
for i in uniqueValues:
    file1.writelines("\n" + i.upper() + " HEADLINES"+"\n")
    for x in df.index:
        if df["category"][x] == i:
            file1.writelines("-"+df["Title"][x]+"["+df["source"][x]+"]"+"\n")
        else:
            pass
        if x >= df.size:
            file1.writelines("-"*30+"\n")
        else:
            pass

#adding the article text
for i in uniqueValues:
    file1.writelines("-"*30+"\n")
    file1.writelines("\n"+i.upper()+ " HEADLINES"+"\n\n")
    file1.writelines("-"*30+"\n")
    file1.writelines("-"*30+"\n")
    for x in df.index:
        if df["category"][x] == i:
            file1.writelines("\n"+"Date: "+ df["Date"][x].strftime("%B"+" "+"%d" +","+ "%Y" +"\n")
            + "Title: "+df["Title"][x]+"\n"
            +"Authors: " + df["Authors"][x]+"\n"
            +"Source: "+df["source"][x]+"\n\n"
            + df["Text"][x]+"\n\n"
            + df["url"][x] + "\n\n"
            + "-"*30+"\n")
        else:
            pass

print("File is at: " + os.path.realpath("./EarlyBird.txt"))
#closing the text file
file1.close()

print("Done!")
