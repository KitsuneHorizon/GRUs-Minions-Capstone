import pandas as pd
from googletrans import Translator

file_location = r'C:\Users\bbgri\OneDrive\Documents\Mason\DAEN-690-Aug2024\Innovation_list_search_results7.xlsx'
file_output = r'C:\Users\bbgri\OneDrive\Documents\Mason\DAEN-690-Aug2024\Output7.xlsx'

df = pd.read_excel(file_location)

df = df.drop(df[df.Snippet == 'No snippet available'].index)


def translate_df(df, column_name):
    translator = Translator()

    def translate_text(text):
        try:
            translation = translator.translate(text, src='zh-cn', dest='en')
            return translation.text
        except Exception as e:
            print(f"Translation error: {e}")
            return text

    df[column_name] = df[column_name].apply(translate_text)
    return df


df = translate_df(df, 'Query')
df = translate_df(df, 'Title')
df = translate_df(df, 'Snippet')
df.to_excel(file_output)