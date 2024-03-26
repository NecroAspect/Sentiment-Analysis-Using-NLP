import nltk
nltk.download('punkt')
nltk.download('wordnet')
nltk.download('cmudict')

import pandas as pd
import requests
import os
import string
from newspaper import Article
from nltk.tokenize import word_tokenize
from nltk.tokenize import sent_tokenize
from nltk.stem import WordNetLemmatizer
from nltk.corpus import cmudict
import re

os.chdir("<parent directory path>")

def extract_article_text(url):
    try:
        article = Article(url)
        article.download()
        article.parse()

        article_title = article.title
        article_text = article.text

        return article_title, article_text
    except Exception as e:
        print(f"Error occurred while extracting text from {url}: {e}")
        return None, None
    



input_file = "Input.xlsx"
df = pd.read_excel(input_file)

output_dir = "./extracted_articles"



for index, row in df.iterrows():
    url_id = row['URL_ID']
    url = row['URL']

    article_title, article_text = extract_article_text(url)

    if article_title and article_text:
        # Remove lines starting with "Blackcoffer Insights"
        cleaned_lines = [line for line in article_text.split('\n') if not line.startswith("Blackcoffer Insights")]
        cleaned_text = '\n'.join(cleaned_lines)

        filename = os.path.join(output_dir, f"{url_id}.txt")
        with open(filename, 'w', encoding='utf-8') as file:
            file.write(f"Title: {article_title}\n\n")
            file.write(cleaned_text)
        print(f"Article text extracted from {url} and saved to {filename}")
    else:
        print(f"Failed to extract text from {url}")

print("Extraction complete.\n\n\n")




folder_path = "./StopWords"

files = []
for filename in os.listdir(folder_path):
    file_path = folder_path + "/" + filename
    file_ptr = open(file_path , "r" , encoding = "latin-1")
    files.append(file_ptr)




all_stop_words = []
for file_ptr in files:
    file_content = file_ptr.read()
    for i in (word_tokenize(file_content)):
      all_stop_words.append(i)





while "|" in all_stop_words:
  all_stop_words.remove("|")

lemmatizer = WordNetLemmatizer()
all_stop_words = [lemmatizer.lemmatize(token) for token in all_stop_words]



positive_word_list = []
negative_word_list = []

MasterDictionary_path = "./MasterDictionary"
positive_file_ptr = open("./MasterDictionary/positive-words.txt" , "r", encoding = "latin-1")
negative_file_ptr = open("./MasterDictionary/negative-words.txt" , "r", encoding = "latin-1")

positive_file_content = positive_file_ptr.read()
negative_file_content = negative_file_ptr.read()

positive_words = word_tokenize(positive_file_content)
negative_words = word_tokenize(negative_file_content)




variables = ["POSITIVE SCORE" , "NEGATIVE SCORE" , "POLARITY SCORE" , "SUBJECTIVITY SCORE" , "AVG SENTENCE LENGTH" , "PERCENTAGE OF COMPLEX WORDS" , "FOG INDEX" , "AVG NUMBER OF WORDS PER SENTENCE" , "COMPLEX WORD COUNT" , "WORD COUNT" ,	"SYLLABLE PER WORD" ,	"PERSONAL PRONOUNS" ,	"AVG WORD LENGTH"]
op_df = pd.read_excel("./Output Data Structure empty.xlsx")




articles_directory_path = "./extracted_articles"
values = {}

for filename in os.listdir(articles_directory_path):
    values[f"{filename[:-4]}"] = {}

    positive_score = 0
    negative_score = 0

    file_path = os.path.join(articles_directory_path , filename)
    if os.path.isfile(file_path):
        filename_index = filename[:-4]
        file_ptr = open(file_path , "r", encoding="utf8")
        file_content = file_ptr.read()
        text_to_tokens = word_tokenize(file_content)
        lemmatizer = WordNetLemmatizer()
        cleaned_tokens = [lemmatizer.lemmatize(token) for token in text_to_tokens]


        cleaned_tokens = [word for word in text_to_tokens if word not in all_stop_words]
        for word in positive_words :
            positive_score += cleaned_tokens.count(word)
        for word in negative_words :
            negative_score += cleaned_tokens.count(word)

        # print(f"for file {filename_index}, positive score = {positive_score} and negative score = {negative_score}")
        values[filename_index]["POSITIVE SCORE"] = positive_score
        values[filename_index]["NEGATIVE SCORE"] = negative_score




d = cmudict.dict()
def num_syll(word):
    try:
        return len([y for y in d[word.lower()][0] if y[-1].isdigit()])
    except KeyError:
        return 0
    




for filename in os.listdir(articles_directory_path):
    file_path = os.path.join(articles_directory_path , filename)

    if os.path.isfile(file_path):
        filename_index = filename[:-4]
        file_ptr = open(file_path , "r", encoding="utf8")
        file_content = file_ptr.read()
        text_to_tokens = word_tokenize(file_content)

        cleaned_tokens = [word for word in text_to_tokens if word not in all_stop_words]

        pos_score = values[filename_index]['POSITIVE SCORE']
        neg_score = values[filename_index]['NEGATIVE SCORE']
        polarity_score = (pos_score - neg_score)/((pos_score + neg_score) + 0.000001)
        subjectivity_score = (pos_score + neg_score)/(len(cleaned_tokens) + 0.000001)

        sentences = sent_tokenize(file_content)
        avg_sentence_len = len(text_to_tokens)/len(sentences);

        complex_count = 0
        for word in text_to_tokens:
            if num_syll(word) > 2:
                complex_count += 1
        complex_percentage = (complex_count)/len(text_to_tokens)

        fog_index = 0.4*(avg_sentence_len + complex_percentage)

        avg_words_per_sentence = len(text_to_tokens)/len(sentences)


        word_count = len(cleaned_tokens)

        total_syllables = 0
        for word in text_to_tokens:
            total_syllables += num_syll(word)
        avg_syll_per_word = total_syllables/len(text_to_tokens)

        pattern = r"\b(I|we|my|ours|us)\b"
        matches = re.findall(pattern, file_content, flags=re.IGNORECASE)
        count_personal_pronouns = len(matches)

        total_charas = 0
        for word in text_to_tokens:
            total_charas += len(word)
        avg_word_length = total_charas/len(text_to_tokens)




        values[filename_index]["POLARITY SCORE"] = polarity_score
        values[filename_index]["SUBJECTIVITY SCORE"] = subjectivity_score
        values[filename_index][ "AVG SENTENCE LENGTH"] = avg_sentence_len
        values[filename_index]["PERCENTAGE OF COMPLEX WORDS"] = complex_percentage
        values[filename_index]["FOG INDEX"] = fog_index
        values[filename_index]["AVG NUMBER OF WORDS PER SENTENCE"] = avg_words_per_sentence
        values[filename_index]["COMPLEX WORD COUNT"] = complex_count
        values[filename_index]["WORD COUNT"] = word_count
        values[filename_index]["SYLLABLE PER WORD"] = avg_syll_per_word
        values[filename_index]["PERSONAL PRONOUNS"] = count_personal_pronouns
        values[filename_index]["AVG WORD LENGTH"] = avg_word_length





op_df = pd.read_excel("./Output Data Structure empty.xlsx")

for (key, value_container) in values.items():
    for prop,value in value_container.items():
        op_df.loc[op_df["URL_ID"] == key, prop] = value

op_df.to_excel("./Output Data Structure.xlsx", index=False)




# Free the file pointers used
file.close()
file_ptr.close()