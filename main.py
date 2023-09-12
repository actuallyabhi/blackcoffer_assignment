import os
import re
import requests
import nltk
import pandas as pd
from bs4 import BeautifulSoup
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize

####################
# Base Setup
####################
nltk.download('punkt')
nltk.download('stopwords')
stop_words = set(stopwords.words('english'))

# Load the input Excel file
input_file = "Input.xlsx"
df = pd.read_excel(input_file)

# Create a directory to store the text files
if not os.path.exists("TextFiles"):
    os.makedirs("TextFiles")

####################
# functions for various tasks
####################


def scrape_data(url):
    try:
        response = requests.get(url)
        if response.status_code == 200:
            soup = BeautifulSoup(response.text, 'html.parser')
            content = ''
            content += soup.find('h1').get_text()
            content += soup.find(class_='td-post-content').get_text()
            return content
        else:
            return None
    except Exception as e:
        print(f"Error scraping {url}: {e}")
        return None

####################
# functions for various calculations
####################


def count_syllables(word):
    exceptions = ["es", "ed"]

    # Remove common word endings
    for exception in exceptions:
        if word.endswith(exception):
            word = word[:-len(exception)]

    # Count the number of vowels in the word
    vowels = "AEIOUaeiou"
    syllable_count = sum(1 for char in word if char in vowels)

    # Ensure that there is at least one syllable, even if no vowels are present
    if syllable_count == 0:
        syllable_count = 1

    return syllable_count


def getAvgWordLength(text):
    # Tokenize the text into words
    words = re.findall(r'\b\w+\b', text)

    # Calculate the total number of characters in all words
    total_characters = sum(len(word) for word in words)

    # Calculate the total number of words
    total_words = len(words)

    # Calculate the average word length
    average_word_length = total_characters / total_words

    return average_word_length


def getPersonalPronounsCount(text):
    pattern = r'\b(?:I|we|my|ours|us)\b'

    # Find all matches of the pattern in the text
    matches = re.findall(pattern, text, flags=re.IGNORECASE)

    # Filter out occurrences of "US" (country name)
    filtered_matches = [match for match in matches if match.lower() != "us"]

    # Count the total number of personal pronouns
    personal_pronoun_count = len(filtered_matches)

    return personal_pronoun_count


def getPoleScores(words):
    positive_score = 0
    negative_score = 0
    for word in words:
        if word in positive_words:
            positive_score += 1
        elif word in negative_words:
            negative_score += 1

    negative_score *= -1
    return positive_score, negative_score


####################
# Main program
####################

print("Starting to scrape content from URLs...")
# Iterate through the rows of the DataFrame
success_count = 0
error_count = 0
for index, row in df.iterrows():
    url_id = str(row["URL_ID"])
    url = str(row["URL"])

    # Scrape the data
    text_content = scrape_data(url)
    print(f"{index+1}. Scraping {url}...")
    if text_content is not None:
        # Save the text file
        output_file = os.path.join("TextFiles", f"{url_id}.txt")
        # handle if file already exists
        if os.path.exists(output_file):
            with open(output_file, "a", encoding="utf-8") as text_file:
                text_file.write(text_content)
        else:
            with open(output_file, "w", encoding="utf-8") as text_file:
                text_file.write(text_content)
        success_count += 1
        print(f"File saved to {output_file}")
    else:
        error_count += 1
        print("No data scraped")

print("Done scarping content from URLs")
print(f"Total URLs: {len(df)}")
print(f"Successful: {success_count}")
print(f"Errors: {error_count}")

####################
# Analyzing the text files
####################

print("Starting to analyze the text files...")
stop_words_lists = []
stop_words_files = ['StopWords_Auditor.txt', 'StopWords_Auditor.txt', 'StopWords_DatesandNumbers.txt',
                    'StopWords_Generic.txt', 'StopWords_GenericLong.txt', 'StopWords_Geographic.txt', 'StopWords_Names.txt']

for file_path in stop_words_files:
    path = os.path.join('WordLists/StopWords', file_path)
    with open(path, 'r') as file:
        stop_words = {line.strip() for line in file}
        stop_words_lists.append(stop_words)

stop_words = set.union(*stop_words_lists)
stop_words = list(stop_words)

# Get word lists
with open("WordLists/MasterDictionary/positive-words.txt", 'r', encoding="utf-8") as file:
    positive_words = file.read().splitlines()
with open("WordLists/MasterDictionary/negative-words.txt", 'r', encoding="ISO-8859-1") as file:
    negative_words = file.read().splitlines()


analyzed_data = []

for index, row in df.iterrows():
    url_id = str(row["URL_ID"])
    url = str(row["URL"])

    # Load the text file
    text_file = "TextFiles/" + url_id + ".txt"
    print(f"{index+1}. Analyzing {text_file}...")
    try:
        with open(text_file, 'r') as file:
            text = file.read()
            sentences = re.split(r'[.!?]', text)

            personal_pronoun_count = getPersonalPronounsCount(text)
            average_word_length = getAvgWordLength(text)
            total_words = 0

            # Tokenize and count words in each sentence
            for sentence in sentences:
                words = re.findall(r'\b\w+\b', sentence)
                total_words += len(words)

            # Calculate the average number of words per sentence
            average_words_per_sentence = total_words / len(sentences)

            # word count
            word_count = 0
            # complex word count
            complex_word_count = 0
            total_syllables = 0

            stop_words = set(stopwords.words('english'))

            for sentence in sentences:
                # Tokenize the sentence into words
                words = word_tokenize(sentence)

                for word in words:
                    # Check if the word is not a stop word and is not punctuation
                    if word not in stop_words and word.isalpha():
                        word_count += 1
                        # Count syllables in the word
                        syllable_count = count_syllables(word)
                        total_syllables += syllable_count
                        if syllable_count > 2:
                            complex_word_count += 1

            # Calculate Average Syllables per Word
            average_syllables_per_word = total_syllables / word_count

            # Calculate Average Sentence Length
            average_sentence_length = word_count / len(sentences)

            # Calculate Percentage of Complex Words
            percentage_complex_words = (complex_word_count / word_count) * 100

            # Calculate Fog Index
            fog_index = 0.4 * (average_sentence_length +
                               percentage_complex_words)

            text = text.replace('\n', ' ')
            words = nltk.word_tokenize(text)
            filtered_words = [word for word in words if all(
                word.lower() not in stop_word for stop_word in stop_words)]

            # Get positive and negative word counts
            positive_score, negative_score = getPoleScores(filtered_words)

            # Calculate Polarity Score
            polarity_score = (positive_score - negative_score) / \
                ((positive_score + negative_score) + 0.000001)

            # Calculate Subjectivity Score
            subjectivity_score = (
                positive_score + negative_score) / ((word_count) + 0.000001)
                
            # append to analyzed_data
            analyzed_data.append([url_id,
                                  url,
                                  positive_score,
                                  negative_score,
                                  polarity_score,
                                  subjectivity_score,
                                  average_sentence_length,
                                  percentage_complex_words,
                                  fog_index,
                                  average_words_per_sentence,
                                  complex_word_count,
                                  word_count,
                                  average_syllables_per_word,
                                  personal_pronoun_count,
                                  average_word_length,
                                  ])

    except FileNotFoundError:
        print("File not found: " + text_file)
        analyzed_data.append(
            [url_id, url, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0])
        continue

####################
# Saving results to an Output.xlsx file
####################

# Create a Pandas DataFrame from the list
df = pd.DataFrame(analyzed_data, columns=["URL_ID", "URL", "POSITIVE SCORE", "NEGATIVE SCORE", "POLARITY SCORE", "SUBJECTIVITY SCORE", "AVG SENTENCE LENGTH",
                  "PERCENTAGE OF COMPLEX WORDS", "FOG INDEX", "AVG NUMBER OF WORDS PER SENTENCE", "COMPLEX WORD COUNT", "WORD COUNT", "SYLLABLE PER WORD", "PERSONAL PRONOUNS", "AVG WORD LENGTH"])

# Save the DataFrame to an Excel file
output_file = "Output.xlsx"
df.to_excel(output_file, index=False)
print("Analysis complete. Output saved to " + output_file)
