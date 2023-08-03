import os
import nltk
import pandas as pd
from nltk.tokenize import word_tokenize
import string

# Function to read stop words from a file


def read_stop_words(file_path):
    with open(file_path, 'r') as file:
        stop_words = [word.strip() for word in file.readlines()]
    return set(stop_words)

# Function to count the number of complex words in the text


def count_complex_words(text):
    # Tokenize the text into words
    words = word_tokenize(text)

    # Count the number of complex words (words with more than two syllables)
    complex_words = [word for word in words if count_syllables(word) > 2]

    return len(complex_words)

# Function to count the total cleaned words in the text


def count_total_cleaned_words(text, stop_words):
    # Tokenize the text into words
    words = word_tokenize(text)

    # Remove stop words and punctuations
    cleaned_words = [word.lower() for word in words if word.lower(
    ) not in stop_words and word.lower() not in string.punctuation]

    return len(cleaned_words)

# Function to calculate the number of syllables in a word


def count_syllables(word):
    vowels = "aeiouy"
    count = 0
    word = word.lower()
    if word[0] in vowels:
        count += 1
    for index in range(1, len(word)):
        if word[index] in vowels and word[index - 1] not in vowels:
            count += 1
    if word.endswith('e'):
        count -= 1
    if count == 0:
        count = 1
    return count


# Create an empty list to store the complex word count and total cleaned word count
results = []

# Load stop words from the stop_words folder
stop_words_folder = "StopWords"
stop_words = set()
for filename in os.listdir(stop_words_folder):
    file_path = os.path.join(stop_words_folder, filename)
    stop_words |= read_stop_words(file_path)

# Process text files one by one from the "extracted_articles" folder
extracted_articles_folder = "extracted_articles"
for filename in os.listdir(extracted_articles_folder):
    file_path = os.path.join(extracted_articles_folder, filename)
    with open(file_path, 'r', encoding='utf-8') as file:
        text = file.read()

    # Count the number of complex words and total cleaned words in each text file
    complex_word_count = count_complex_words(text)
    total_cleaned_words = count_total_cleaned_words(text, stop_words)

    # Store the results in a dictionary
    result_dict = {
        # Remove '.txt' extension from the filename
        'Title': os.path.splitext(filename)[0],
        'Complex Word Count': complex_word_count,
        'Total Cleaned Word Count': total_cleaned_words
    }

    # Append the result dictionary to the list
    results.append(result_dict)

# Convert the list of dictionaries to a DataFrame
df = pd.DataFrame(results)

# Save the DataFrame to an Excel file
output_file = 'complex_word_and_total_word_count.xlsx'
df.to_excel(output_file, index=False)

print(
    f"Complex word count and total cleaned word count results saved to {output_file}")
