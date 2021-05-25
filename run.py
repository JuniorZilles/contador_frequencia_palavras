
from nltk.corpus import stopwords 
from nltk.tokenize import word_tokenize 
from nltk import  FreqDist, download
import pandas as pd
import os

def getPath(pasta:str)->list:
    lista = []
    for nome in os.listdir(pasta):
        lista.append({"caminho": os.path.join(pasta, nome), "nome":nome.replace('.xls', "")})
    return lista

def downloadStopWords()->None:
    download('stopwords')
    download('punkt')

def replaceItems(value:str)->str:
    return value.replace(':', '').replace('-', '').replace('.', '').replace('(', '').replace(')', '').replace(',', '').replace('+', '')

def removeStopWords(samples:str)->str:
    stop_wordsPort = set(stopwords.words('portuguese')) 
    stop_wordsIngl = set(stopwords.words('english')) 
    word_tokens = word_tokenize(replaceItems(samples)) 
    filtered_sentence = [w for w in word_tokens if not w in stop_wordsPort and not w in stop_wordsIngl] 
    return filtered_sentence

def calculaFrequencia(samples:str) -> dict:
    return FreqDist(samples)

def main():
    folder = 'xlsx'
    paths = getPath(folder)
    listaConteudo = []
    #downloadStopWords()
    for path in paths:
        df = pd.read_excel(path['caminho'])
        df.columns = list(range(29))
        for row in df.index:
            titulo = df[7][row].lower()
            listaConteudo += removeStopWords(titulo)
        freq = calculaFrequencia(listaConteudo)
        sort = sorted(freq.items(), key=lambda item: item[1], reverse=True)
        with open('frequencias.txt', 'w', encoding='UTF-8') as wr:
            for a in sort:
                wr.write(f'{a[0]}: {a[1]} \n')
main()