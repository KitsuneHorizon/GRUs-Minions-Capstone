"""HELP FUNCTION:
Required Input: Text file

NOTE: DO NOT USE A TEXT FILE WITH MANDARIN. ENGLISH ONLY FOR THIS SET OF CODE
"""


import matplotlib.pyplot as plt
from wordcloud import WordCloud
import nltk
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize
import string

# Download stopwords if not already downloaded
#nltk.download('punkt')
#nltk.download('stopwords')

def generate_word_cloud_from_file():
    # User input for the text file
    file_path = input("Enter the path and text file name: ")

    try:
        # Open and read the file
        with open(file_path, 'r') as file:
            text = file.read()

        # Tokenize words and remove punctuation
        tokens = word_tokenize(text.lower())
        words = [word for word in tokens if word.isalpha()]

        # Remove stopwords
        stop_words = set(stopwords.words('english'))
        filtered_words = [word for word in words if word not in stop_words]

        # Create word cloud
        wordcloud = WordCloud(width=800, height=400, background_color='white').generate(' '.join(filtered_words))

        # Display the word cloud
        plt.figure(figsize=(10, 5))
        plt.imshow(wordcloud, interpolation='bilinear')
        plt.axis('off')
        plt.show()

    except FileNotFoundError:
        print("The file was not found. Please check the path and try again.")

# Call the function
generate_word_cloud_from_file()
