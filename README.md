# Product Matching Engine

A sophisticated web application designed to match products between internal catalogs and customer lists using advanced similarity algorithms and threshold-based matching.

## 🎯 Project Overview

The Product Matching Engine allows users to upload two product datasets (internal catalog and customer list) and find potential matches based on configurable similarity thresholds. The application uses a combination of TF-IDF vectorization and fuzzy string matching to generate confidence scores for product pairs.

## ✨ Key Features

- **Threshold-Based Matching**: Configurable similarity threshold (0-100%) to filter matches
- **Dual Algorithm Approach**: Combines TF-IDF cosine similarity and fuzzy string matching
- **Data Standardization**: Automatic text normalization, unit conversion, and stop-word removal
- **Clean Web Interface**: Streamlit-based UI with intuitive file upload and results display
- **Standalone Executable**: Packaged as a single .exe file for easy distribution

## 🚀 Quick Start

### Prerequisites
- Python 3.8 or higher
- Required packages listed in `requirements.txt`

### Installation

1. Clone the repository:
```bash
git clone <repository-url>
cd CodeOutTool/product-matching-engine
```

2. Create and activate conda environment:
```bash
conda create -n product-matcher python=3.9
conda activate product-matcher
```

3. Install dependencies:
```bash
pip install -r requirements.txt
```

### Running the Application

#### Method 1: Streamlit App
```bash
streamlit run app.py
```

#### Method 2: Executable
Run the provided `ProductMatcher.exe` (if available) or build it using:
```bash
pyinstaller --name "ProductMatcher" --onefile --windowed app.py
```

## 📁 Project Structure

```
product-matching-engine/
├── app.py                    # Main Streamlit application
├── src/                      # Source code modules
│   ├── config.py            # Configuration settings
│   ├── processing.py        # Data processing logic
│   └── gtin_processing.py   # GTIN-specific processing
├── requirements.txt          # Python dependencies
├── README.md                # This file
├── samples/                 # Sample data files
└── build_scripts/          # Build and packaging scripts
```

## 🎮 Usage Guide

1. **Upload Files**: 
   - Upload your internal product catalog (CSV/Excel)
   - Upload the customer's product list (CSV/Excel)

2. **Configure Threshold**: 
   - Set the similarity threshold (default: 80%)
   - Higher thresholds = more strict matching

3. **Find Matches**: 
   - Click "Find Matches" to process the data
   - Results will display grouped by customer product

4. **Review Results**: 
   - Each match shows confidence score
   - Results filtered by your threshold setting

## 🔧 Technical Implementation

### Data Processing Pipeline

1. **Text Normalization**: Lowercase conversion, punctuation removal
2. **Unit Standardization**: Convert measurements to base units (grams, milliliters)
3. **Stop Word Removal**: Eliminate non-descriptive words
4. **Vectorization**: TF-IDF for keyword significance
5. **Similarity Scoring**: Combined TF-IDF + fuzzy matching

### Matching Algorithm

The application uses a weighted combination of:
- **TF-IDF Cosine Similarity** (50% weight): Identifies significant shared keywords
- **Fuzzy String Matching** (50% weight): Handles word order and token differences

Final score calculation:
```
combined_score = (tfidf_score * 0.5) + (fuzzy_score * 0.5)
```

## 📊 Supported File Formats

- **CSV**: Comma-separated values
- **Excel**: .xlsx files
- **Expected Columns**: 
  - `product_name` (required)
  - `description` (optional)
  - `size` (optional, for unit processing)

## 🛠️ Development

### Adding New Features

1. Modify the core processing logic in `src/processing.py`
2. Update the UI components in `app.py`
3. Test with sample data files

### Testing

Run the test suite:
```bash
python test_size_matching.py
```

### Building Executable

Use the provided build scripts:
```bash
# Standard build
build.bat

# Debug build
build_debug.bat
```

## 📈 Performance Considerations

- **Large Datasets**: Processing time scales with O(n*m) where n and m are dataset sizes
- **Memory Usage**: TF-IDF vectorization requires memory proportional to vocabulary size
- **Threshold Impact**: Higher thresholds reduce processing time and result size

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests for new functionality
5. Submit a pull request

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 🆘 Support

For issues and questions:
1. Check the troubleshooting section in `CONTRIBUTING.md`
2. Review sample data files for expected format
3. Test with smaller datasets first

## 🔄 Version History

- **v1.0.0**: Initial release with core matching functionality
- **v1.1.0**: Added unit conversion and size processing
- **v1.2.0**: Enhanced UI and error handling

---

**Built with ❤️ using Streamlit, scikit-learn, and thefuzz**
