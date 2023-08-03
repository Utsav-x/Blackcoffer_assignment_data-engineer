import os
import pandas as pd
import nltk
from nltk.tokenize import word_tokenize
import re

# Function to count the number of syllables in a word


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
    if word.endswith(('es', 'ed')) and len(word) > 2:
        count -= 1
    if count == 0:
        count = 1
    return count

# Function to count the number of syllables per word in the text


def count_syllables_per_word(text):
    # Tokenize the text into words
    words = word_tokenize(text)

    # Count the number of syllables in each word
    syllables_per_word = [count_syllables(word) for word in words]

    return syllables_per_word

# Function to count personal pronouns in the text


def count_personal_pronouns(text):
    # Define the regex pattern to find personal pronouns
    pronoun_pattern = r'\b(I|we|my|ours|us)\b'

    # Find all matches of personal pronouns using regex
    personal_pronouns = re.findall(pronoun_pattern, text, re.IGNORECASE)

    # Filter out the country name "US"
    personal_pronouns = [
        pronoun for pronoun in personal_pronouns if pronoun.lower() != 'us']

    return len(personal_pronouns)

# Function to calculate the average word length


def calculate_avg_word_length(text):
    # Tokenize the text into words
    words = word_tokenize(text)

    # Calculate the sum of the total number of characters in each word
    total_characters = sum(len(word) for word in words)

    # Calculate the total number of words
    total_words = len(words)

    # Calculate the average word length
    avg_word_length = total_characters / total_words

    return avg_word_length


# Create an empty list to store the analysis results
results = []

# Process text files one by one from the "extracted_articles" folder
extracted_articles_folder = "extracted_articles"
for filename in os.listdir(extracted_articles_folder):
    file_path = os.path.join(extracted_articles_folder, filename)
    with open(file_path, 'r', encoding='utf-8') as file:
        text = file.read()

    # Calculate various metrics for each text file
    syllables_per_word = count_syllables_per_word(text)
    personal_pronouns_count = count_personal_pronouns(text)
    avg_word_length = calculate_avg_word_length(text)

    # Store the results in a dictionary
    result_dict = {
        # Remove '.txt' extension from the filename
        'Title': os.path.splitext(filename)[0],
        'Syllables Per Word': syllables_per_word,
        'Personal Pronouns Count': personal_pronouns_count,
        'Average Word Length': avg_word_length
    }

    # Append the result dictionary to the list
    results.append(result_dict)

# Convert the list of dictionaries to a DataFrame
df = pd.DataFrame(results)

# Save the DataFrame to an Excel file
output_file = 'syllable_personal_average_results.xlsx'
df.to_excel(output_file, index=False)

print(f"Text analysis results saved to {output_file}")
