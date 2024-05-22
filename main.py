import pandas as pd
import requests
from bs4 import BeautifulSoup
import nltk
from nltk.tokenize import word_tokenize, sent_tokenize
from nltk.corpus import cmudict
from nltk.corpus import stopwords
import re


# Download NLTK resources
nltk.download('punkt')
nltk.download('averaged_perceptron_tagger')
nltk.download('cmudict')

# Function to extract article text from URL
def extract_article_text(url):
    response = requests.get(url)
    soup = BeautifulSoup(response.content, 'html.parser')
    
    # Extracting article text
    article_text = ''
    article_body = soup.find('body')
    for paragraph in article_body.find_all('p'):
        article_text += paragraph.text.strip() + '\n'
    
    return article_text

# Load input data
input_data = pd.read_excel('input.xlsx')

# Function to compute variables
def calculate_sentiment(text):
    # Tokenize text into sentences and words
    sentences = sent_tokenize(text)
    words = word_tokenize(text)
    
    

def calculate_readability(text):
    sentences = sent_tokenize(text)
    words = word_tokenize(text)
    
    # Average sentence length
    avg_sentence_length = len(words) / len(sentences)
    
    # Percentage of complex words
    complex_words = [word for word in words if len(word) > 2]
    percentage_complex_words = (len(complex_words) / len(words)) * 100
    
    # Fog Index
    fog_index = 0.4 * (avg_sentence_length + percentage_complex_words)
    
    return avg_sentence_length, percentage_complex_words, fog_index


# Function to compute variables
def compute_variables(text):
    # Tokenize text
    words = word_tokenize(text)
    sentences = sent_tokenize(text)
    
    # Counting number of words
    word_count = len(words)
    
    # Counting number of sentences
    sentence_count = len(sentences)
    
    # Average number of words per sentence
    avg_words_per_sentence = word_count / sentence_count
    
    # Calculating syllables per word
    d = cmudict.dict()
    syllable_count = sum([len([y for y in x if y[-1].isdigit()]) for word in words for x in d.get(word.lower(), [[]])])
    syllable_per_word = syllable_count / word_count
    
    # Other variables to compute
    # You can compute other variables as specified in the assignment
    
    return word_count, avg_words_per_sentence, syllable_per_word

def average_words_per_sentence(text):
    sentences = sent_tokenize(text)
    total_words = len(word_tokenize(text))
    average_words_per_sentence = total_words / len(sentences)
    return average_words_per_sentence

def count_complex_words(text):
    words = word_tokenize(text)
    complex_words = [word for word in words if len(word) > 2]
    return len(complex_words)

def count_total_words(text):
    words = word_tokenize(text)
    total_words = len([word for word in words if word.isalnum()])
    return total_words

def count_syllables(word):
    vowels = 'aeiou'
    count = 0
    prev_char = ""
    for char in word.lower():
        if char in vowels and prev_char != char:
            count += 1
        prev_char = char
    
    # Handle exceptions for words ending with "es" or "ed"
    if word.lower().endswith(('es', 'ed')) and count > 1:
        count -= 1 
    return count
    

def calculate_syllable_per_word(text):
    words = word_tokenize(text)
    total_syllables = sum(count_syllables(word) for word in words)
    total_words = len(words)
    return total_syllables / total_words

def count_personal_pronouns(text):
    personal_pronouns = ['i', 'we', 'my', 'ours', 'us']
    # Exclude 'US' if it's not a personal pronoun
    text = re.sub(r'\bUS\b', '', text)
    pronoun_count = sum(1 for word in word_tokenize(text) if word.lower() in personal_pronouns)
    return pronoun_count

def calculate_average_word_length(text):
    words = word_tokenize(text)
    total_characters = sum(len(word) for word in words)
    total_words = len(words)
    return total_characters / total_words


# Extracting computing variables for each URL
output_data = []

for index, row in input_data.iterrows():
    url = row['URL']
    url_id = row['URL_ID']
    article_text = extract_article_text(url)

    # Readability Analysis
    avg_sentence_length, percentage_complex_words, fog_index = calculate_readability(article_text)

    # Other Metrics
    avg_words_per_sentence = average_words_per_sentence(article_text)
    complex_word_count = count_complex_words(article_text)
    total_word_count = count_total_words(article_text)
    syllable_per_word = calculate_syllable_per_word(article_text)
    personal_pronouns_count = count_personal_pronouns(article_text)
    average_word_length = calculate_average_word_length(article_text)
    word_count, avg_words_per_sentence, syllable_per_word = compute_variables(article_text)
    
    # Append extracted data to output list
    output_data.append({
        'URL_ID': url_id,
        'URL': url,
        'Average Sentence Length': [avg_sentence_length],
        'Percentage of Complex Words': [percentage_complex_words],
        'Fog Index': [fog_index],
        'Average Words Per Sentence': [avg_words_per_sentence],
        'Complex Word Count': [complex_word_count],
        'Total Word Count': [total_word_count],
        'Syllable Per Word': [syllable_per_word],
        'Personal Pronouns Count': [personal_pronouns_count],
        'Average Word Length': [average_word_length]
    })

# Create DataFrame from output data
output_df = pd.DataFrame(output_data)

# Save output DataFrame to Excel file
output_df.to_excel('Output.xlsx', index=False)

