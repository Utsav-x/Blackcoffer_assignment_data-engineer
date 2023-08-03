import os
import nltk
import pandas as pd
from nltk.tokenize import sent_tokenize, word_tokenize

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

# Function to calculate the Fog Index, average sentence length, and percentage of complex words


def analyze_readability(text):
    # Tokenize the text into sentences
    sentences = sent_tokenize(text)

    # Tokenize the sentences into words
    words = word_tokenize(text)

    # Calculate the average sentence length
    avg_sentence_length = len(words) / len(sentences)

    # Calculate the number of complex words
    complex_words = [word for word in words if count_syllables(word) >= 3]
    percentage_complex_words = len(complex_words) / len(words)

    # Calculate the Fog Index
    fog_index = 0.4 * (avg_sentence_length + percentage_complex_words)

    return fog_index, avg_sentence_length, percentage_complex_words


# Create an empty list to store the readability analysis results
readability_results = []

# Process text files one by one from the "extracted_articles" folder
extracted_articles_folder = "extracted_articles"
for filename in os.listdir(extracted_articles_folder):
    file_path = os.path.join(extracted_articles_folder, filename)
    with open(file_path, 'r', encoding='utf-8') as file:
        text = file.read()

    # Perform the readability analysis for each text file
    fog_index, avg_sentence_length, percentage_complex_words = analyze_readability(
        text)

    # Store the readability analysis results in a dictionary
    result_dict = {
        # Remove '.txt' extension from the filename
        'Title': os.path.splitext(filename)[0],
        'Fog Index': fog_index,
        'Average Sentence Length': avg_sentence_length,
        'Percentage of Complex Words': percentage_complex_words
    }

    # Append the readability analysis result dictionary to the list
    readability_results.append(result_dict)

# Convert the list of readability analysis dictionaries to a DataFrame
readability_df = pd.DataFrame(readability_results)

# Save the readability analysis DataFrame to an Excel file
output_file = 'readability_analysis_results.xlsx'
readability_df.to_excel(output_file, index=False)

print(f"Readability analysis results saved to {output_file}")
