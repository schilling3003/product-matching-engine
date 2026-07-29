import io
import pandas as pd
import streamlit as st
from sklearn.feature_extraction.text import TfidfVectorizer

from src.processing import clean_and_standardize


def _file_hash(file):
    """Stable hash for uploaded files using Streamlit's file_id."""
    if hasattr(file, "file_id"):
        return file.file_id
    # Fallback for non-Streamlit file-like objects
    pos = file.tell()
    file.seek(0)
    data = file.read()
    file.seek(pos)
    return hash(data)


@st.cache_data(show_spinner=False, hash_funcs={io.BytesIO: _file_hash})
def load_uploaded_file(file):
    """Load an uploaded CSV or Excel file once and cache it for reruns."""
    if file is None:
        return None
    file.seek(0)
    if file.name.lower().endswith(".csv") or "csv" in file.name.lower():
        return pd.read_csv(file, low_memory=False)
    return pd.read_excel(file)


@st.cache_data(show_spinner=False)
def get_cleaned_df(df, column_config, remove_stop_words, case_sensitive, include_size_in_text):
    """Clean and standardize a DataFrame; cache by inputs."""
    return clean_and_standardize(df, column_config, remove_stop_words, case_sensitive, include_size_in_text)


@st.cache_data(show_spinner=False)
def get_vectorizer_and_vectors(cleaned_catalog_df, cleaned_customer_df, remove_stop_words, enable_text_matching):
    """Build TF-IDF vectorizer and sparse matrices; cache by cleaned inputs."""
    if not enable_text_matching:
        return None, None, None
    vectorizer = TfidfVectorizer(stop_words="english" if remove_stop_words else None)
    all_texts = pd.concat(
        [cleaned_catalog_df["combined_product_name"], cleaned_customer_df["combined_product_name"]]
    ).dropna()
    vectorizer.fit(all_texts)
    catalog_vectors = vectorizer.transform(cleaned_catalog_df["combined_product_name"].fillna(""))
    customer_vectors = vectorizer.transform(cleaned_customer_df["combined_product_name"].fillna(""))
    return vectorizer, catalog_vectors, customer_vectors
