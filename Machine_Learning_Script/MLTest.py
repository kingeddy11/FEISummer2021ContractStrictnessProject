import spacy
import string
from spacy.lang.en import English
import pandas as pd
import numpy as np
import glob
from spacy import displacy
from spacy.lang.en.stop_words import STOP_WORDS
from sklearn.model_selection import train_test_split
from sklearn.feature_extraction.text import CountVectorizer, TfidfVectorizer
from sklearn.linear_model import LogisticRegression
from sklearn.pipeline import Pipeline
from sklearn.metrics import accuracy_score, confusion_matrix, classification_report
from sklearn.feature_extraction.text import TfidfTransformer
from sklearn.model_selection import KFold

#reading in datasets
path = r'C:\Users\edwar\OneDrive\Documents\FEISummer2021\CreditContractStrictnessProfBattaSum2021\Scraping' # use your path
csv_files = glob.glob(path + "/*.csv")

li = []

for filename in csv_files:
    df = pd.read_csv(filename, index_col=None, header=0)
    li.append(df)


Book_Net_Worth_def = li[0]
Capex_and_acqui = li[1]
Cash_Equivalents = li[2]
Consolidated_Interest_Coverage_Ratio =  li[3]
Continuity_of_Operations = li[4]
Covenant_Components = li[5]
Disposals = li[6]
Dynamic_Thresholds = li[7]
EBITDA = li[8]
EBITDA = EBITDA.dropna(subset = ['Text', 'General category'])
EBITDA = EBITDA.reset_index()
Financial_Covenants = li[9]
Fixed_Charges = li[10]
Indebtedness = li[11]
Indebtedness_def = li[12]
Interest_Charges = li[13]
Investments = li[14]
Investments_and_Restricted_Payments = li[15]
Liens = li[16]
Lien_def = li[17]
Negative_Covenants = li[18]
Net_Income = li[19]
Net_Income = Net_Income.dropna(subset = ['Text', 'General category'])
Net_Income = Net_Income.reset_index()
# Net_Income = Net_Income.reindex(columns = ['IntelligizeID', 'Item', 'Specific Definition', 'Text', 'General category'])
Net_Operating_Income =  li[20]
Other_def = li[21]
Permitted_Investment = li[22]
Permitted_Liens = li[23]
Permitted_Transfer = li[24]
Restricted_Payments = li[25]
Restrictions_Lessee = li[26]
Subordinated_Debt = li[27]
Tangible_Net_Worth = li[28]
Total_Debt = li[29]
Valuation = li[30]

#Removing NA values in Text and General Category column

#loading nlp object for text
nlp = spacy.load('en_core_web_sm')
text = EBITDA['Text'][1]


#  "nlp" Object is used to create documents with linguistic annotations.
my_doc = nlp(text)

# Create our list of punctuation marks
punctuations = string.punctuation

# Create our list of stopwords
stop_words = STOP_WORDS

# Create list of original word tokens
token_list = []
for token in my_doc:
    token_list.append(token.text)

print(token_list)

# Creating our tokenizer function
def spacy_tokenizer(text):
    # Creating our token object, which is used to create documents with linguistic annotations.
    token = nlp(text)

    # Lemmatizing each token and coverting each token into lowercase
    token = [ word.lemma_.lower().strip() if word.lemma_ != "-PRON-" else word.lower_ for word in token ]

    # Removing stop words
    token = [ word for word in token if word not in stop_words and word not in punctuations ]

    #Converting token list object to string object
    # token = ' '.join(map(str, token))
    # return preprocessed string of token
    return token


#Applying Tokenizer Function on Text column of Datasets
# EBITDA['Text'] = EBITDA['Text'].apply(spacy_tokenizer)
# #print(EBITDA['Text'])

# #word count for data
# EBITDA['Text'].apply(lambda x: len(x.split(' '))).sum()

tfidf = TfidfVectorizer(tokenizer =  spacy_tokenizer)
classifier =  LogisticRegression()

#EBITDA
X_EBITDA = EBITDA['Text']
y_EBITDA = EBITDA['General category']
# X_train_EBITDA, X_test_EBITDA, y_train_EBITDA, y_test_EBITDA = train_test_split(X_EBITDA, y_EBITDA, test_size=0.3, random_state = 42)

# 5-fold cross validation
kf = KFold(n_splits=5, shuffle=True, random_state=42)
for train_index, test_index in kf.split(X_EBITDA):
    X_train_EBITDA, X_test_EBITDA = X_EBITDA[train_index], X_EBITDA[test_index]
    y_train_EBITDA, y_test_EBITDA = y_EBITDA[train_index], y_EBITDA[test_index]

logreg = Pipeline([('tfidf', tfidf),
                ('clf', classifier),
               ])

logreg.fit(X_train_EBITDA, y_train_EBITDA)

y_pred_EBITDA = logreg.predict(X_test_EBITDA)

print('accuracy %s' % accuracy_score(y_pred_EBITDA, y_test_EBITDA))
print(classification_report(y_test_EBITDA, y_pred_EBITDA))
print(confusion_matrix(y_test_EBITDA, y_pred_EBITDA))

#Net Income
X_Net_Income = Net_Income['Text']
y_Net_Income = Net_Income['General category']
# X_train_Net_Income, X_test_Net_Income, y_train_Net_Income, y_test_Net_Income = train_test_split(X_Net_Income, y_Net_Income, test_size=0.3, random_state = 42)

#5-fold cross validation
kf = KFold(n_splits=5, shuffle=True, random_state=42)
for train_index, test_index in kf.split(X_Net_Income):
    X_train_Net_Income, X_test_Net_Income = X_Net_Income[train_index], X_Net_Income[test_index]
    y_train_Net_Income, y_test_Net_Income = y_Net_Income[train_index], y_Net_Income[test_index]

logreg = Pipeline([('tfidf', tfidf),
                ('clf', classifier),
               ])

logreg.fit(X_train_Net_Income, y_train_Net_Income)

y_pred_Net_Income = logreg.predict(X_test_Net_Income)

print('accuracy %s' % accuracy_score(y_pred_Net_Income, y_test_Net_Income))
print(classification_report(y_test_Net_Income, y_pred_Net_Income))
print(confusion_matrix(y_test_Net_Income, y_pred_Net_Income))
