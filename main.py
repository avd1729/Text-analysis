import requests
from bs4 import BeautifulSoup
from textblob import TextBlob
import nltk
from nltk.corpus import cmudict
import pandas as pd
import re

# Load the syllable dictionary
d = cmudict.dict()

# Function to count syllables in a word


def syllable_count(word):
    try:
        return [len(list(y for y in x if y[-1].isdigit())) for x in d[word.lower()]][0]
    except KeyError:
        return len(re.findall(r'[aeiouy]+', word.lower()))

# Text analysis functions


def analyze_text(text):
    blob = TextBlob(text)

    # Positive and Negative Score
    positive_score = sum(1 for word in blob.words if TextBlob(
        word).sentiment.polarity > 0)
    negative_score = sum(1 for word in blob.words if TextBlob(
        word).sentiment.polarity < 0)

    # Polarity and Subjectivity Score
    polarity_score = blob.sentiment.polarity
    subjectivity_score = blob.sentiment.subjectivity

    # Sentence statistics
    sentences = blob.sentences
    num_sentences = len(sentences)
    words = blob.words
    num_words = len(words)

    avg_sentence_length = num_words / num_sentences if num_sentences else 0

    # Complex words and Fog Index
    complex_words = [word for word in words if syllable_count(word) >= 3]
    percentage_of_complex_words = len(
        complex_words) / num_words if num_words else 0
    fog_index = 0.4 * (avg_sentence_length + percentage_of_complex_words * 100)

    # Avg number of words per sentence
    avg_number_of_words_per_sentence = avg_sentence_length

    # Complex word count
    complex_word_count = len(complex_words)

    # Syllables per word
    syllables_per_word = sum(syllable_count(word)
                             for word in words) / num_words if num_words else 0

    # Personal pronouns
    personal_pronouns = len(re.findall(r'\b(I|we|my|ours|us)\b', text, re.I))

    # Avg word length
    avg_word_length = sum(len(word) for word in words) / \
        num_words if num_words else 0

    return {
        'POSITIVE SCORE': positive_score,
        'NEGATIVE SCORE': negative_score,
        'POLARITY SCORE': polarity_score,
        'SUBJECTIVITY SCORE': subjectivity_score,
        'AVG SENTENCE LENGTH': avg_sentence_length,
        'PERCENTAGE OF COMPLEX WORDS': percentage_of_complex_words * 100,
        'FOG INDEX': fog_index,
        'AVG NUMBER OF WORDS PER SENTENCE': avg_number_of_words_per_sentence,
        'COMPLEX WORD COUNT': complex_word_count,
        'WORD COUNT': num_words,
        'SYLLABLE PER WORD': syllables_per_word,
        'PERSONAL PRONOUNS': personal_pronouns,
        'AVG WORD LENGTH': avg_word_length
    }

# Function to scrape and analyze text from a URL


def analyze_url(url):
    response = requests.get(url)
    soup = BeautifulSoup(response.content, 'html.parser')

    # Extract text from paragraphs
    text = ' '.join(p.get_text() for p in soup.find_all('p'))

    return analyze_text(text)


# Read URLs from the Excel file
input_file_path = 'C:/Users/Aravind/PROJECTS/Text-analysis/Data/Input.xlsx'
df = pd.read_excel(input_file_path)

# Assuming URLs are in the first column
urls = df.iloc[:, 1]

# Analyze each URL and store results
count = 1
results = []
for url in urls:
    try:
        result = analyze_url(url)
        results.append(result)
        print(f"Successfully scraped {count} articles")
        count += 1
    except Exception as e:
        print(f"Error processing {url}: {e}")

# Convert results to a DataFrame
results_df = pd.DataFrame(results)

# Save the results to a new Excel file
output_file_path = 'C:/Users/Aravind/PROJECTS/Text-analysis/Data/Text_Analysis_Results.xlsx'
results_df.to_excel(output_file_path, index=False)

print(f"Analysis complete. Results saved to {output_file_path}.")
