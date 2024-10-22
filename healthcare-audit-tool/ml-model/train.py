import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.naive_bayes import MultinomialNB
from sklearn.pipeline import make_pipeline
import pickle

# Load training data from database (pseudo code)
data = fetch_training_data()

# Preprocess data
texts = data['content']
labels = data['classification']

# Create a pipeline that vectorizes the text data and then applies Naive Bayes
model = make_pipeline(TfidfVectorizer(), MultinomialNB())

# Train the model
model.fit(texts, labels)

# Save the model
with open('ml-model/model.pkl', 'wb') as f:
    pickle.dump(model, f)
