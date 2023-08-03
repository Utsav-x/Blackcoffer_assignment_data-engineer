import os
import pandas as pd
import requests
from bs4 import BeautifulSoup


def extract_article_text(url, url_id):
    try:
        response = requests.get(url)
        response.raise_for_status()

        soup = BeautifulSoup(response.content, 'html.parser')

        # Remove unwanted elements like header, footer, and others
        for element in soup(['header', 'footer', 'nav', 'script', 'style']):
            element.extract()

        # Extract article title and text
        article_title = soup.title.text.strip()
        article_text = '\n'.join([p.get_text() for p in soup.find_all('p')])

        return article_title, article_text
    except Exception as e:
        print(f"Error while extracting from URL_ID {url_id}: {e}")
        return None, None


def main():
    input_file = "input.xlsx"
    output_file = "extracted_articles.xlsx"

    # Read the input Excel file
    df = pd.read_excel(input_file)

    # Create empty lists to store the extracted data
    extracted_titles = []
    extracted_texts = []

    for index, row in df.iterrows():
        url_id = row['URL_ID']
        url = row['URL']

        article_title, article_text = extract_article_text(url, url_id)

        if article_title and article_text:
            extracted_titles.append(article_title)
            extracted_texts.append(article_text)
            print(f"Extracted article from URL_ID {url_id}")
        else:
            print(f"Failed to extract article from URL_ID {url_id}")

    # Create a new DataFrame to store the extracted data
    extracted_data = pd.DataFrame({
        'Title': extracted_titles,
        'Text': extracted_texts
    })

    # Save the extracted data to a new Excel file
    extracted_data.to_excel(output_file, index=False)

    print(f"Extracted data saved to {output_file}")


if __name__ == "__main__":
    main()
