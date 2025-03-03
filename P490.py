# P490 Resume Classification Project.
#==============================================================================

# Importing the libraries
import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import pdfplumber
from docx import Document
from PyPDF2 import PdfReader
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.layout import LAParams
from pdfminer.converter import TextConverter
from io import StringIO
from pdfminer.pdfpage import PDFPage
from docx import Document
import comtypes.client

import warnings
warnings.filterwarnings('ignore')
pd.set_option('display.max_columns', None)
#------------------------------------------------------------------------------
# Data Preprocessing

# Function to read text from pdf files
def extract_text_from_pdf(path_to_pdf):
    resource_manager = PDFResourceManager(caching=True)
    out_text = StringIO()
    laParams = LAParams()
    text_converter = TextConverter(resource_manager, out_text, laparams=laParams)
    fp = open(path_to_pdf, 'rb')
    interpreter = PDFPageInterpreter(resource_manager, text_converter)

    for page in PDFPage.get_pages(fp, pagenos=set(), maxpages=0, password="", caching=True, check_extractable=True):
        interpreter.process_page(page)

    text = out_text.getvalue()
    fp.close()
    text_converter.close()
    out_text.close()
    return text

# Function to read text from docx files
def extract_text_from_docx(path_to_file):
    doc_object = open(path_to_file, "rb")
    doc_reader = Document(doc_object)
    data = ""
    for p in doc_reader.paragraphs:
        data += p.text + "\n"
    return data

# Function to read text from doc files
def convert_doc_to_docx(doc_path):
    word = comtypes.client.CreateObject('Word.Application')
    doc = word.Documents.Open(doc_path)
    docx_path = doc_path + "x"
    doc.SaveAs(docx_path, FileFormat=16)  # 16 represents the docx format
    doc.Close()
    word.Quit()
    return docx_path

# Writing a function to extract text from all the file formats and  storing in a dataframe

def extract_text_from_files_in_directories(directories):
    extracted_data = []

    for category, directory_path in directories.items():
        for filename in os.listdir(directory_path):
            file_path = os.path.join(directory_path, filename)

            if filename.lower().endswith('.pdf'):
                text = extract_text_from_pdf(file_path)
                extracted_data.append({"Content": category, "Extracted Information": text})

            elif filename.lower().endswith('.docx'):
                text = extract_text_from_docx(file_path)
                extracted_data.append({"Content": category, "Extracted Information": text})

            elif filename.lower().endswith('.doc'):
                docx_path = convert_doc_to_docx(file_path)
                text = extract_text_from_docx(docx_path)
                extracted_data.append({"Content": category, "Extracted Information": text})

    # Convert the list of dictionaries to a DataFrame
    df = pd.DataFrame(extracted_data)
    return df



# Directories and categories
# Usage Example:
directories = {
    'PeopleSoft': r'D:\ExcelR\Data_Science_Projects\P490_ResumeAnalysis\Resumes\Peoplesoft resumes',
    'React JS Developer': r'D:\ExcelR\Data_Science_Projects\P490_ResumeAnalysis\Resumes',
    'SQL Developer': r'D:\ExcelR\Data_Science_Projects\P490_ResumeAnalysis\Resumes\SQL Developer Lightning insight',
    'Workday': r'D:\ExcelR\Data_Science_Projects\P490_ResumeAnalysis\Resumes\workday resumes'
}

df = extract_text_from_files_in_directories(directories)
print(df)

csv_file_path = r'D:\ExcelR\Data_Science_Projects\P490_ResumeAnalysis\extracted1_data.csv'
df.to_csv(csv_file_path, index=False, encoding='utf-8')
print(f"Data saved to {csv_file_path}")


#==============================================================================
# Reading the Saved CSV file using pandas function
# Data Preprocessing

resume_df=pd.read_csv('extracted1_data.csv')


resume_df.head()

resume_df.shape

# Checking and Handling Null Values
resume_df.isna().sum()

# Checking and Handling Duplicate Records
resume_df.duplicated().sum()

resume_df.drop_duplicates(inplace=True)


resume_df.duplicated().sum()
resume_df.shape

# Checking the unique values in the category column
resume_df['Content'].unique()

# Value Counts of Category Columns
category_counts=resume_df['Content'].value_counts()
category_counts

#------------------------------------------------------------------------------
# EDA

# Visualizations

plt.figure(figsize=(10,6))
sns.barplot(x=resume_df['Content'].index,y=resume_df['Content'].values,palette='viridis')
plt.xticks(rotation=45)
plt.title('Distribution of Target Variable')
plt.xlabel('Category')
plt.ylabel('Count')
plt.show()



# Counting the words and character counts and Visualizing them
resume_df['word_count']=resume_df['Extracted Information'].apply(lambda x:len(str(x).split()))
resume_df['Char_count']=resume_df['Extracted Information'].apply(lambda x:len(str(x)))

# Visualizations of Word Count & Character Count
plt.figure(figsize=(10,6))
sns.histplot(resume_df['word_count'],kde=True,bins=30)
plt.title('Distribution of Word Count')
plt.xlabel('Word Count')
plt.show()

# Visualizations of Character Count
plt.figure(figsize=(10,6))
sns.histplot(resume_df['Char_count'],kde=True,bins=30)
plt.title('Distribution of Character Count')
plt.xlabel('Character Count')
plt.show()

# Plotting a Word Cloud to find the most frequent words before cleaning the text
from wordcloud import WordCloud
text=''.join(resume_df['Extracted Information'])
wordcloud=WordCloud(width=800,height=400,background_color='white').generate(text)

plt.figure(figsize=(10,6))
plt.imshow(wordcloud,interpolation='bilinear')
plt.axis('off')
plt.title('Most frequest words in the resumes')
plt.show()



# Finding the top 20 words and visualizing them before removing the stop words
from collections import Counter
all_words=''.join(resume_df['Extracted Information']).split()
common_words=Counter(all_words).most_common(20)

words,counts=zip(*common_words)
plt.figure(figsize=(10,6))
sns.barplot(x=counts,y=words,palette='coolwarm')
plt.title('Top 20 Common Words')
plt.xlabel('Frequency')
plt.ylabel('Words')
plt.show()



# Plotting a Pie Chart
plt.figure(figsize=(10,8))
plt.pie(category_counts,labels=category_counts.index,autopct='%1.1f%%',startangle=140,colors=plt.cm.Paired.colors)
plt.legend(category_counts.index, loc='upper left', bbox_to_anchor=(1, 0.5))
plt.title('Category Distribution in Resumes')
plt.show()



# Word Distribution by Category

plt.figure(figsize=(12,8))
sns.boxplot(data=resume_df,x='Content',y='word_count',palette='pastel')
plt.xticks(rotation=45)
plt.title('Word Count Distribution by category')
plt.xlabel('Category')
plt.ylabel('Word count')
plt.show()



# Plotting Word Count using Violin plot

plt.figure(figsize=(12,8))
sns.violinplot(data=resume_df,x='Content',y='word_count',palette='muted')
plt.xticks(rotation=45)
plt.title('Word Count Distribution using Violin plot by Category')
plt.xlabel('Category')
plt.ylabel('Word Count')
plt.show()


# Category Wise Finding and plotting the top 20 words.
for category1 in resume_df['Content'].unique():
    # Filter the resumes for the current category
    category_resumes = resume_df[resume_df['Content'] == category1]['Extracted Information']
    
    # Combine all text for the current category
    category_text = ' '.join(category_resumes)
    
    # Split text into words
    category_words = category_text.split()
    
    # Count word frequencies
    word_counts = Counter(category_words)
    
    # Get the top 20 words
    top_20_words = word_counts.most_common(20)
    words1, counts1 = zip(*top_20_words)
    
    # Plot the top 20 words for the category
    plt.figure(figsize=(10, 6))
    sns.barplot(x=counts1, y=words1, palette='viridis')
    plt.title(f'Top 20 Words in {category1} Resumes')
    plt.xlabel('Frequency')
    plt.ylabel('Words')
    plt.tight_layout()
    plt.show()

#==============================================================================


resume_df['Extracted Information'][0]
 

# Text Cleaning

import re
import nltk
from nltk.tokenize import word_tokenize
from nltk.corpus import stopwords
from nltk.stem import WordNetLemmatizer

nltk.download('punkt')
nltk.download('stopwords')
nltk.download('wordnet')



# Initialize Lemmatizer and Stopwords
lemmatizer = WordNetLemmatizer()
stop_words = set(stopwords.words('english'))

# Function to clean text
def clean_text(text):
    text = text.lower()  # Convert to lowercase
    text = re.sub(r'http\S+|www\S+', '', text)  # Remove URLs
    text = re.sub(r'[^a-zA-Z0-9\s]', '', text)  # Remove punctuation
    text = re.sub(r'\s+', ' ', text).strip()  # Remove extra spaces
    
    words = word_tokenize(text)  # Tokenization
    words = [word for word in words if word not in stop_words]  # Remove stopwords
    words = [lemmatizer.lemmatize(word) for word in words]  # Lemmatization
    
    return ' '.join(words)  # Reconstruct text

# Apply cleaning to the 'Resume_Text' column
resume_df['Cleaned_Resume_Text'] = resume_df['Extracted Information'].apply(clean_text)



# Plotting a Word Cloud to find the most frequent words after text cleaning
text1=''.join(resume_df['Cleaned_Resume_Text'])
wordcloud1=WordCloud(width=800,height=400,background_color='white').generate(text1)

plt.figure(figsize=(10,6))
plt.imshow(wordcloud1,interpolation='bilinear')
plt.axis('off')
plt.title('Most frequest words in the resumes after text cleaning')
plt.show()


# Plotting a bar graph to find the top 20 words after removing stop words


all_text2 = ' '.join(resume_df['Cleaned_Resume_Text'])
word_counts2 = Counter(all_text2.split())
top_words2 = word_counts2.most_common(20)
words2, counts2 = zip(*top_words2)

# Plot the top 20 words
plt.figure(figsize=(10, 6))
plt.barh(words2, counts2, color='skyblue')
plt.xlabel("Frequency")
plt.ylabel("Words")
plt.title("Top 20 Words in Resumes")
plt.gca().invert_yaxis()  
plt.show()

resume_df.head()

resume_df['Cleaned_Resume_Text'].isna().sum()
resume_df.duplicated().sum()

#------------------------------------------------------------------------------
# Label Encoding
# Converting the category column values into numeric values using label Encoding method
from sklearn.preprocessing import LabelEncoder

le=LabelEncoder()
resume_df['Content']=le.fit_transform(resume_df['Content'])

resume_df.shape
#------------------------------------------------------------------------------

# Feature Engineering
from sklearn.feature_extraction.text import TfidfVectorizer


vectorizer = TfidfVectorizer(sublinear_tf=True,stop_words='english')

tfidf_matrix = vectorizer.fit_transform(resume_df['Cleaned_Resume_Text'])

df_tfidf=pd.DataFrame(tfidf_matrix.toarray(),columns=vectorizer.get_feature_names_out())

# df_tfidf.shape

# df_tfidf.duplicated().sum()

# df_tfidf.drop_duplicates(inplace=True)

# df_tfidf.duplicated().sum()
#------------------------------------------------------------------------------

# Splitting the Data into training and testing

from sklearn.model_selection import train_test_split

x=df_tfidf
y=resume_df['Content']


X_train,X_test,y_train,y_test=train_test_split(x,y,test_size=0.2,random_state=42)

X_train.shape # (63, 4156)
X_test.shape  # (16, 4156)

y_train.shape # (63,)
y_test.shape  # (16,)

#------------------------------------------------------------------------------

# Model Building
from sklearn.ensemble import RandomForestClassifier, GradientBoostingClassifier
from sklearn.naive_bayes import MultinomialNB
from sklearn.svm import SVC
from sklearn.metrics import accuracy_score, confusion_matrix, classification_report
from sklearn.tree import DecisionTreeClassifier
from sklearn.linear_model import LogisticRegression

# Model 1: Naive Baye's Classifier
#----------------------------------

# Model 1 : Naive Bayes
nb_model = MultinomialNB()
nb_model.fit(X_train, y_train)
y_pred_nb = nb_model.predict(X_test)
accuracy_nb = accuracy_score(y_test, y_pred_nb)
print("Naive Bayes Accuracy:", accuracy_nb)

# Confusion Matrix for Naive Bayes
plt.figure(figsize=(8, 6))
cm_nb = confusion_matrix(y_test, y_pred_nb)
sns.heatmap(cm_nb, annot=True, fmt='d', cmap='Blues')
plt.title('Confusion Matrix - Naive Bayes')
plt.xlabel('Predicted')
plt.ylabel('Actual')
plt.show()


nb_train=nb_model.score(X_train, y_train)
print('Training Accuracy:',nb_train)


nb_test=nb_model.score(X_test, y_test)
print('Testing Accuracy:',nb_test)

from sklearn.model_selection import cross_val_score

cv_scores_nb = cross_val_score(nb_model, X_train, y_train, cv=5, scoring='accuracy')
print("Cross-Validated Accuracy:",cv_scores_nb)
print("Mean cross-validation accuracy:", cv_scores_nb.mean())

#-----------------------------------------------------------------------------

# Model 2: Random Forest Classifier
#----------------------------------

rf_model = RandomForestClassifier(random_state=42)
rf_model.fit(X_train, y_train)
y_pred_rf = rf_model.predict(X_test)
accuracy_rf = accuracy_score(y_test, y_pred_rf)
print("Random Forest Accuracy:", accuracy_rf)

# Confusion Matrix for Random Forest
plt.figure(figsize=(8, 6))
cm_rf = confusion_matrix(y_test, y_pred_rf)
sns.heatmap(cm_rf, annot=True, fmt='d', cmap='Blues')
plt.title('Confusion Matrix - Random Forest')
plt.xlabel('Predicted')
plt.ylabel('Actual')
plt.show()



rfc_train=rf_model.score(X_train, y_train)
print('Training Accuracy:',rfc_train)


rfc_test=rf_model.score(X_test, y_test)
print('Testing Accuracy:',rfc_test)

cv_scores_rf = cross_val_score(rf_model, X_train, y_train, cv=5, scoring='accuracy')
print("Cross-Validated Accuracy:",cv_scores_rf)
print("Mean cross-validation accuracy:", cv_scores_rf.mean())


#-----------------------------------------------------------------------------

# Model 3: Support Vector Machine(SVM)
svm_model = SVC(kernel='linear', random_state=42)
svm_model.fit(X_train, y_train)
y_pred_svm = svm_model.predict(X_test)
accuracy_svm = accuracy_score(y_test, y_pred_svm)
print("SVM Accuracy:", accuracy_svm)

# Confusion Matrix for SVM
plt.figure(figsize=(8, 6))
cm_svm = confusion_matrix(y_test, y_pred_svm)
sns.heatmap(cm_svm, annot=True, fmt='d', cmap='Blues')
plt.title('Confusion Matrix - SVM')
plt.xlabel('Predicted')
plt.ylabel('Actual')
plt.show()


svm_train=svm_model.score(X_train, y_train)
print('Training Accuracy:',svm_train)


svm_test=svm_model.score(X_test, y_test)
print('Testing Accuracy:',svm_test)


cv_scores_svm = cross_val_score(svm_model, X_train, y_train, cv=5, scoring='accuracy')
print("Cross-Validation Scores:", cv_scores_svm)
print("Mean CV Accuracy:", cv_scores_svm.mean())

#-----------------------------------------------------------------------------
# Model 4: Decision Tree

dt_model = DecisionTreeClassifier(random_state=42)
dt_model.fit(X_train, y_train)
y_pred_dt = dt_model.predict(X_test)
accuracy_dt = accuracy_score(y_test, y_pred_dt)
print("Decision Tree Accuracy:", accuracy_dt)

# Confusion Matrix for Decision Tree
plt.figure(figsize=(8, 6))
cm_dt = confusion_matrix(y_test, y_pred_dt)
sns.heatmap(cm_dt, annot=True, fmt='d', cmap='Blues')
plt.title('Confusion Matrix - Decision Tree')
plt.xlabel('Predicted')
plt.ylabel('Actual')
plt.show()


dt_train=dt_model.score(X_train, y_train)
print('Training Accuracy:',dt_train)


dt_test=dt_model.score(X_test, y_test)
print('Testing Accuracy:',dt_test)

cv_scores_dt = cross_val_score(dt_model, X_train, y_train, cv=5, scoring='accuracy')
print("Cross-Validation Scores:", cv_scores_dt)
print("Mean CV Accuracy:", cv_scores_dt.mean())


#------------------------------------------------------------------------------

# Model 5: Gradient Boosting

gb_model = GradientBoostingClassifier(random_state=42)
gb_model.fit(X_train, y_train)
y_pred_gb = gb_model.predict(X_test)
accuracy_gb = accuracy_score(y_test, y_pred_gb)
print("Gradient Boosting Accuracy:", accuracy_gb)

# Confusion Matrix for Gradient Boosting
plt.figure(figsize=(8, 6))
cm_gb = confusion_matrix(y_test, y_pred_gb)
sns.heatmap(cm_gb, annot=True, fmt='d', cmap='Blues')
plt.title('Confusion Matrix - Gradient Boosting')
plt.xlabel('Predicted')
plt.ylabel('Actual')
plt.show()


gb_train=gb_model.score(X_train, y_train)
print('Training Accuracy:',gb_train)


gb_test=gb_model.score(X_test, y_test)
print('Testing Accuracy:',gb_test)

cv_scores_gb = cross_val_score(gb_model, X_train, y_train, cv=5, scoring='accuracy')
print("Cross-Validation Scores:", cv_scores_gb)
print("Mean CV Accuracy:", cv_scores_gb.mean())


#------------------------------------------------------------------------------
# Model 6: KNNClassifier
from sklearn.neighbors import KNeighborsClassifier

knn = KNeighborsClassifier(n_neighbors=5)
knn.fit(X_train, y_train)

y_pred_knn = knn.predict(X_test)

accuracy_knn=accuracy_score(y_test, y_pred_knn)
print("Accuracy Score:", accuracy_knn)


# Confusion Matrix for Gradient Boosting
plt.figure(figsize=(8, 6))
cm_knn = confusion_matrix(y_test, y_pred_knn)
sns.heatmap(cm_knn, annot=True, fmt='d', cmap='Blues')
plt.title('Confusion Matrix - KNN')
plt.xlabel('Predicted')
plt.ylabel('Actual')
plt.show()



knn_train=knn.score(X_train, y_train)
print('Training Accuracy:',knn_train)


knn_test=knn.score(X_test, y_test)
print('Testing Accuracy:',knn_test)


#==============================================================================
# Model Results Comparision
data={'Model':pd.Series(['Naive Bayes','Random Forest','SVM','Decision Tree',
                         'Gradient Boosting','KNNClassifier']),
      'Accuracies':pd.Series([accuracy_nb,accuracy_rf,accuracy_svm,
                              accuracy_dt,accuracy_gb,accuracy_knn])}
table_acc=pd.DataFrame(data)
table_acc.sort_values(['Accuracies'])

#=============================================================================
# Training and Testing accuracies

# Comparing the Results

data1={'Model':pd.Series(['Naive Bayes','Random Forest','SVM','Decision Tree',
                         'Gradient Boosting','KNNClassifier']),
      'Train Accuracies':pd.Series([nb_train,rfc_train,svm_train,
                              dt_train,gb_train,knn_train]),
      'Test Accuracies':pd.Series([nb_test,rfc_test,svm_test,
                              dt_test,gb_test,knn_test])
      
      }
results1=pd.DataFrame(data1)
results1.sort_values(['Test Accuracies'])
#==============================================================================

import pickle
filename = 'modelRF.pkl'
pickle.dump(rf_model,open(filename,'wb'))

filename = 'vector.pkl'
pickle.dump(vectorizer,open(filename,'wb'))

filename='label_encoder.pkl'
pickle.dump(le, open(filename,'wb'))