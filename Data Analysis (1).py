#!/usr/bin/env python
# coding: utf-8

# # Data Extraction

# In[9]:


import requests
from bs4 import BeautifulSoup
import openpyxl
import os

# Load the input.xlsx file
input_file = r"C:\Users\Nidhi Gangwar\Downloads\Input.xlsx"
wb = openpyxl.load_workbook(input_file)
sheet = wb.active

# Create a directory to save the articles
output_dir = 'articles'
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# Function to extract article text from a URL
def extract_article(url):
    try:
        # Send a request to the URL
        response = requests.get(url)
        soup = BeautifulSoup(response.content, 'html.parser')

        # Extract the article title
        title = soup.find('h1').get_text(strip=True)

        # Extract the article text
        paragraphs = soup.find_all('p')
        article_text = '\n'.join([para.get_text(strip=True) for para in paragraphs])

        return title + '\n\n' + article_text
    except Exception as e:
        print(f"Error extracting article from {url}: {e}")
        return None

# Iterate through each row in the Excel sheet
for row in sheet.iter_rows(min_row=2, values_only=True):
    url_id, url = row

    # Extract the article text
    article_content = extract_article(url)

    if article_content:
        # Save the extracted text in a .txt file with URL_ID as the filename
        file_path = os.path.join(output_dir, f'{url_id}.txt')
        with open(file_path, 'w', encoding='utf-8') as file:
            file.write(article_content)

print("Extraction completed!")


# # Data Analysis

# In[10]:


get_ipython().system('pip install nltk')
get_ipython().system('pip install textblob')
get_ipython().system('pip install pandas')


# In[21]:



import os
import pandas as pd
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize, sent_tokenize
from nltk import pos_tag
from textblob import TextBlob
import re

import nltk
nltk.download('punkt')
nltk.download('stopwords')
nltk.download('averaged_perceptron_tagger')

input_file = r"C:\Users\Nidhi Gangwar\Downloads\Input.xlsx"
output_structure_file = r"C:\Users\Nidhi Gangwar\Downloads\Output Data Structure.xlsx"
output_dir ='articles'

with open(r"C:\Users\Nidhi Gangwar\Downloads\positive-words.txt", "r") as file:
    positive_words = set(file.read().split())
with open(r"C:\Users\Nidhi Gangwar\Downloads\negative-words.txt", "r") as file:
    negative_words = set(file.read().split())

def syllable_count(word):
    word = word.lower()
    vowels = "aeiouy"
    count = 0
    if word[0] in vowels:
        count += 1
    for index in range(1, len(word)):
        if word[index] in vowels and word[index - 1] not in vowels:
            count += 1
    if word.endswith("e"):
        count -= 1
    if count == 0:
        count += 1
    return count

def analyze_text(text):
    words = word_tokenize(text)
    sentences = sent_tokenize(text)

    positive_score = sum(1 for word in words if word in positive_words)
    negative_score = sum(1 for word in words if word in negative_words)
    polarity_score = (positive_score - negative_score) / ((positive_score + negative_score) + 0.000001)
    subjectivity_score = TextBlob(text).sentiment.subjectivity

    word_count = len(words)
    complex_words = sum(1 for word in words if syllable_count(word) > 2)
    percentage_complex_words = complex_words / word_count

    avg_sentence_length = word_count / len(sentences)
    avg_words_per_sentence = avg_sentence_length

    fog_index = 0.4 * (avg_sentence_length + percentage_complex_words * 100)
    avg_syllables_per_word = sum(syllable_count(word) for word in words) / word_count
    avg_word_length = sum(len(word) for word in words) / word_count

    personal_pronouns = len([word for word, tag in pos_tag(words) if tag in ['PRP', 'PRP$', 'WP', 'WP$']])

    return {
        "POSITIVE SCORE": positive_score,
        "NEGATIVE SCORE": negative_score,
        "POLARITY SCORE": polarity_score,
        "SUBJECTIVITY SCORE": subjectivity_score,
        "AVG SENTENCE LENGTH": avg_sentence_length,
        "PERCENTAGE OF COMPLEX WORDS": percentage_complex_words,
        "FOG INDEX": fog_index,
        "AVG NUMBER OF WORDS PER SENTENCE": avg_words_per_sentence,
        "COMPLEX WORD COUNT": complex_words,
        "WORD COUNT": word_count,
        "SYLLABLE PER WORD": avg_syllables_per_word,
        "PERSONAL PRONOUNS": personal_pronouns,
        "AVG WORD LENGTH": avg_word_length
    }

input_data = pd.read_excel(input_file)
output_data = pd.read_excel(output_structure_file)

for index, row in input_data.iterrows():
    url_id = row['URL_ID']
    file_path = os.path.join(output_dir, f'{url_id}.txt')

    if os.path.exists(file_path):
        print(f"Processing file: {file_path}")
        with open(file_path, 'r', encoding='utf-8') as file:
            text = file.read()

        analysis_results = analyze_text(text)
        print(f"Results for {url_id}: {analysis_results}")

        for key, value in analysis_results.items():
            output_data.at[index, key] = value
    else:
        print(f"File not found: {file_path}")

output_file_path = r"C:\Users\Nidhi Gangwar\Downloads\Output Data Structure.xlsx" 
output_data.to_excel(output_file_path, index=False)
print(f"Textual analysis completed and results saved to {output_file_path}!")


# In[ ]:




