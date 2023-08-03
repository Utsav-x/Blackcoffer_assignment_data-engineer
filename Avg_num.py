import os
import nltk
import pandas as pd
from nltk.tokenize import sent_tokenize, word_tokenize

# Function to calculate the average number of words per sentence


def calculate_avg_words_per_sentence(text):
    # Tokenize the text into sentences
    sentences = sent_tokenize(text)

    # Tokenize the sentences into words
    words = word_tokenize(text)

    # Calculate the average number of words per sentence
    avg_words_per_sentence = len(words) / len(sentences)

    return avg_words_per_sentence


# Create an empty list to store the average number of words per sentence
avg_num_results = []

# Process text files one by one from the "extracted_articles" folder
extracted_articles_folder = "extracted_articles"
for filename in os.listdir(extracted_articles_folder):
    file_path = os.path.join(extracted_articles_folder, filename)
    with open(file_path, 'r', encoding='utf-8') as file:
        text = file.read()

    # Perform the calculation for each text file
    avg_words_per_sentence = calculate_avg_words_per_sentence(text)

    # Store the results in a dictionary
    result_dict = {
        # Remove '.txt' extension from the filename
        'Title': os.path.splitext(filename)[0],
        'Average Number of Words Per Sentence': avg_words_per_sentence
    }

    # Append the result dictionary to the list
    avg_num_results.append(result_dict)

# Convert the list of dictionaries to a DataFrame
avg_num_df = pd.DataFrame(avg_num_results)

# Save the DataFrame to an Excel file
output_file = 'average_number_of_words_per_sentence.xlsx'
avg_num_df.to_excel(output_file, index=False)

print(f"Average number of words per sentence results saved to {output_file}")
