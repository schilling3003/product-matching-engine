AI Agent Task: Build a Product Matching Engine
1. Project Objective
Your task is to build a web application that serves as a "Product Matching Engine." The application will allow a user to upload two product lists (their internal catalog and a customer's product list) and find potential matches between them.

The core logic will involve data cleaning, standardization, and applying similarity algorithms to generate a confidence score for potential matches.

2. Core Feature Requirement: Threshold-Based Matching
This is the most critical requirement: The application must not simply find the single best match for each product. Instead, it must:

Allow the user to specify a "similarity threshold" (a percentage, e.g., 75%).

For each product in the customer's list, compare it against every product in the internal catalog.

Return a list of all matches from the internal catalog that have a similarity score above the user-defined threshold.

3. Application Workflow & UI Specifications
The user interface should be clean, simple, and guide the user through a step-by-step process.

UI Components:

File Uploaders:

An uploader for "Your Product Catalog" (CSV or Excel).

An uploader for the "Customer's Product List" (CSV or Excel).

The app should expect columns like product_name, description, and size.

Configuration:

A numeric input or slider for the "Similarity Threshold." This should accept values from 0 to 100. Set a default value of 80.

Action Button:

A button labeled "Find Matches". This will trigger the backend matching process.

Results Display:

A clear, tabular display for the output.

The results should be grouped by the "Customer Product."

Columns:

Customer Product Name

Matched Catalog Product Name

Confidence Score (%)

If no matches are found for a customer product above the threshold, it should be indicated (e.g., "No matches found above threshold").

4. Backend Implementation Guide
The backend will contain all the logic for cleaning data and calculating similarity. Use Python with the recommended libraries.

Step 1: Data Cleaning and Standardization Pipeline
This pipeline must be applied to both datasets upon upload. Create a function that performs the following actions on product name and description fields:

Text Normalization: Convert all text to lowercase. Remove all punctuation and extra whitespace.

Unit of Measurement Conversion: Standardize units of measurement to a base unit (e.g., grams for weight, milliliters for volume).

Implement a conversion map: {'oz': 28.35, 'lb': 453.59, 'gallon': 3785.41, ...}.

Parse strings like "16 oz" and convert them to a standardized format like "453.6g".

Stop Word Removal: Remove common, non-descriptive words (e.g., "a", "the", "and", "for", "case", "pack").

Step 2: The Matching Engine
The engine will calculate a combined similarity score.

Keyword Similarity (TF-IDF):

Use scikit-learn's TfidfVectorizer and cosine_similarity.

Fit the vectorizer on the combined corpus of all product names/descriptions from both files to create a shared vocabulary.

This will produce a score based on the significance of shared keywords.

Fuzzy String Matching:

Use thefuzz library.

The fuzz.token_set_ratio function is highly effective as it handles differences in word order and tokenizes the strings.

Step 3: Scoring and Ranking Logic
This is where the core feature requirement is implemented.

Combined Score Function:

Create a function calculate_combined_score(string1, string2) that returns a single score from 0 to 100.

This function should internally calculate both the TF-IDF cosine similarity and the fuzzy string ratio.

Return a weighted average of the two scores. A 50/50 weighting is a good starting point. (tfidf_score * 50) + (fuzzy_score * 50).

Main Processing Loop:

When the user clicks "Find Matches," initiate a process that does the following:

For each product in the cleaned customer list:

Initialize an empty list to store its matches.

For each product in the cleaned internal catalog:

Calculate the combined_score between the customer product and the catalog product.

If combined_score >= user_defined_threshold:

Add the (catalog_product, combined_score) to the matches list for the current customer product.

After iterating, format this data structure for display on the front end.

5. Recommended Technology Stack
Language: Python

Core Libraries:

pandas: For data ingestion and manipulation.

scikit-learn: For TF-IDF implementation.

thefuzz (or fuzzywuzzy): For fuzzy string matching.

Web Framework (Choose one):

Streamlit: Recommended for its speed in creating data-centric applications.

Flask: A good alternative for a more custom web application.

6. Agent Action Plan Summary
Construct the user interface with two file uploaders, a threshold input, a start button, and a results area.

Implement the data cleaning pipeline function.

Implement the calculate_combined_score function using TF-IDF and TheFuzz.

Develop the main application logic that iterates through the product lists and applies the threshold logic.

Ensure the results are displayed clearly in a table, grouped by the original customer product.

Package the application using Streamlit or Flask.

7. Packaging for Distribution
To ensure the application is accessible to non-technical users, it must be packaged as a standalone executable. This eliminates the need for users to install Python, Streamlit, or any dependencies manually.

Tool: Use PyInstaller to bundle the Python application and all its dependencies into a single executable file.

Requirement: The final output should be a single .exe file (for Windows) or a corresponding executable for other operating systems that can be run directly by the user.