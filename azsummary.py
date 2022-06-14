# Module to generate email summary
import os
import openpyxl
import re
import string
from nltk.tokenize import word_tokenize
from nltk.corpus import stopwords

stop_words = set(stopwords.words('english'))
stop_words.update(['not', 'can', 'yall', 'seem', 'surprised', 'bad', 'cannot', 'while', 'says', 'why', 'annoyed', 'big', 'freak', 'ashamed'])

def removeStopWords(text):
    text = re.sub(r'[^\w\s]','',text)
    word_tokens = word_tokenize(text)

    filtered_sentence = []

    for w in word_tokens:
        if w.lower() not in stop_words:
            filtered_sentence.append(w)

    return filtered_sentence

def generateData(excelPath):
    wb = openpyxl.load_workbook(excelPath) 
    wb.active = wb['CS Feedback']
    sheet = wb.active
    data = {}

    issuesCol = sheet['B']
    commentsCol = sheet['F']

    for i in range(len(issuesCol)):
        issuesCell = str(issuesCol[i].value).lower()
        filteredComment = removeStopWords(str(commentsCol[i].value))

        if not issuesCell in data:
            data[issuesCell] = {}

        for word in filteredComment:
            if word.lower() in data[issuesCell]:
                data[issuesCell][word.lower()] += 1
            else:
                data[issuesCell][word.lower()] = 1
    
    wb.close()    
    return data

def wordFrequency(text):
    frequency = {}
    for word in text:
        if word.lower() in frequency:
            frequency[word.lower()] += 1
        else:
            frequency[word.lower()] = 1
    return frequency

def generateIssueSummary(text, prevData):
    text = removeStopWords(text)
    textFrequency = wordFrequency(text)
    wordsInText = set(textFrequency.keys())

    issuesList = list(prevData.keys())
    comparisonWeights = []
    
    for i in range(len(issuesList)):
        issue = issuesList[i]
        comparisonWeight = 0
        
        for keyword in prevData[issue]:
            if keyword in wordsInText:
                comparisonWeight += prevData[issue][keyword] * textFrequency[keyword]
        
        comparisonWeights.append(comparisonWeight)

    print(comparisonWeights)

    maxIndex = 0
    maxValue = 0
    for i in range(len(comparisonWeights)):
        if comparisonWeights[i] > maxValue:
            maxValue = comparisonWeights[i]
            maxIndex = i

    return issuesList[maxIndex]