# Product Matching Engine

This user-friendly application helps you find matching products between two lists (your catalog and a customer's list). **Optimized for performance and accuracy**, it intelligently handles variations in product names, sizes, and even messy barcode (GTIN/UPC) data. Simply upload your files, choose your matching preferences, and let the app do the work.

## Features

### ✅ Core Functionality
- **File Upload**: Supports CSV and Excel files for both product lists.
- **Flexible Column Mapping**: Intelligently detects and configures columns for product names, descriptions, sizes, and GTINs.
- **Threshold-Based Results**: Returns all matches above a configurable similarity threshold.
- **Sample Data**: Includes test files for immediate functionality testing.

### ⚙️ Powerful Matching Methods
- **Text Matching**: Uses a hybrid approach combining semantic understanding (TF-IDF) and spelling similarity (Fuzzy Matching) to find matches even with different phrasing.
- **GTIN Matching**: Robustly matches products using GTIN/UPC/Barcode data. It automatically handles messy data, including missing check digits, mixed formats (GTIN-8/12/13/14), and case-vs-unit relationships.
- **Size Matching**: Compares standardized product sizes (e.g., '12 oz' vs '355ml') with a configurable tolerance, adding another layer of accuracy.
- **Match Restrictions**: Restrict matches to products within the same category, commodity, or other custom columns (available in "Find Similar Within File" mode).

### 🚀 Performance Optimizations
- **Vectorized Operations**: Uses NumPy for high-speed, matrix-based calculations, avoiding slow row-by-row processing.
- **Multiprocessing**: Leverages multiple CPU cores to accelerate CPU-intensive fuzzy string matching on large datasets.
- **Smart Filtering**: Dramatically speeds up processing by intelligently filtering out obviously poor matches early in the pipeline.

### 📊 Enhanced UI & Results
- **Simple & Advanced UI**: An intuitive interface for all users, with an expandable "Advanced Settings" panel for expert control.
- **Comprehensive Statistics**: Get a clear summary of match rates, total matches, and average confidence scores.
- **GTIN Quality Report**: An expandable report that provides detailed statistics on the quality of your GTIN data, including valid, correctable, and invalid codes.
- **Detailed Results Table**: View all matches with sortable columns for overall confidence, TF-IDF, Fuzzy, and GTIN scores.
- **Download Options**: Export your matched results to either CSV or Excel for further analysis.
- **Threshold Explorer**: For within-file grouping, explore how groups change across different similarity thresholds.
- **Excel Integration**: Export enhanced data for Excel-based threshold analysis using Power Pivot or formulas.

## How to Run the Application

### 1. Prerequisites
- Python 3.7+ installed on your system.

### 2. Setup
1.  **Clone the repository or download the source code.**
2.  **Navigate to the project directory:**
    ```bash
    cd product-matching-engine
    ```
3.  **Create and activate a Python virtual environment:**
    - On Windows:
      ```bash
      python -m venv .venv
      .venv\Scripts\activate
      ```
    - On macOS/Linux:
      ```bash
      python3 -m venv .venv
      source .venv/bin/activate
      ```
4.  **Install the required dependencies:**
    ```bash
    pip install -r requirements.txt
    ```

### 3. Running the App
Once the setup is complete, run the application using Streamlit:
```bash
streamlit run app.py
```
This will open the application in your default web browser.

## How to Use

1.  **Upload Files**: Upload your product catalog and customer list (CSV or Excel), or a single file for "Find Similar Within File" mode.
2.  **Configure Columns**: The app will try to auto-detect your columns. Review and adjust the mappings for product names, sizes, and GTINs as needed.
3.  **Adjust Settings**: Use the simple sidebar controls to set your desired matching strictness and enable/disable Text, GTIN, or Size matching.
4.  **(Optional) Set Restrictions**: In "Find Similar Within File" mode, you can restrict matches to products within the same category, commodity, or other columns.
5.  **Find Matches**: Click the "Find Matches" button to start the analysis.
6.  **Review Results**: Examine the match summary, GTIN quality report, and the detailed results table. Download the results if needed.

## Excel Integration

The application offers powerful Excel export capabilities for advanced threshold analysis:

### Standard Export
- **CSV/Excel**: Download basic match results for immediate use.

### Enhanced Threshold Analysis Export
When using "Find Similar Within File" with grouping enabled:
- **Threshold Explorer Workbook**: Contains pre-calculated groups at different thresholds.
- **Enhanced Excel for Analysis**: Includes similarity matrix data for custom Excel analysis.

### Excel Analysis Options
1. **Power Pivot Method** (Recommended): Create dynamic threshold explorer with slicers
2. **Formula Method**: Use Excel formulas for custom threshold analysis
3. **Pivot Table Method**: Analyze patterns with Excel PivotTables

For detailed instructions, see:
- `EXCEL_THRESHOLD_EXPLORER.md` - Complete guide for Excel-based analysis
- `POWER_PIVOT_TEMPLATE.md` - Advanced Power Pivot techniques and templates

## Matching Engine Explained

The application uses a sophisticated, multi-layered approach to find the most accurate matches.

### Matching Methods
- **Text Matching (TF-IDF + Fuzzy)**: This is the default method. It first understands the *meaning* of product names (semantic search) and then checks for *spelling similarities* (fuzzy search). You can control the balance between these two.
- **GTIN Matching**: When enabled, this method looks for matches in your barcode data. It's highly accurate and can find matches even if the GTINs are incomplete or in different formats. If a GTIN match is found, it is weighted heavily in the final score.
- **Size Matching**: This optional method gives a bonus to products with similar sizes, helping to distinguish between, for example, a 1-liter bottle and a 2-liter bottle of the same drink.
- **Match Restrictions**: Available in "Find Similar Within File" mode, this feature allows you to restrict matches to products that share the same values in selected columns (e.g., only match products within the same category).

### Scoring Logic
- The final **Confidence Score** is a weighted average of the enabled matching methods.
- In the default **combined mode** (Text + GTIN), the engine dynamically adjusts weights. If a strong GTIN match is found for a pair of products, the GTIN score is given higher importance. If no GTIN match is found, the score relies entirely on the text and size similarity.
- You can switch to **Text-only** or **GTIN-only** modes for more specific use cases.

## Deployment

### Option 1: Run from Source (Recommended)
Follow the `How to Run the Application` steps above to run the app in a local development environment.

### Option 2: Standalone Executable
An executable can be created using PyInstaller. However, due to known compatibility issues between PyInstaller and recent versions of Streamlit, this method may be unreliable. For the most stable experience, running from the source is recommended.

To build the executable, you can use the provided batch scripts (`build.bat`, `build_fixed.bat`, etc.), but success may vary depending on your system configuration.

## File Format Requirements

- **File Types**: CSV (.csv) and Excel (.xlsx, .xls).
- **Required Columns**: At least one column containing product names/descriptions must be selected for each file.
- **Optional Columns**: For best results, include columns for:
  - **Size**: Can be in a combined format (e.g., "12 oz") or in separate value/unit columns.
  - **GTIN/UPC**: Can be in any standard format (GTIN-8, 12, 13, 14). You can select multiple GTIN columns.
