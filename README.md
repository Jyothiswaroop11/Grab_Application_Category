# 🎯 Applipedia Category Scraper

## 📋 Description
A Python-based web scraping tool that automatically extracts category and subcategory information from Palo Alto Networks Applipedia for given applications or URLs.

## 📈 Version History

### Version 3.0.0 (Current)
- Added color-coded output formatting
  - Green: Successfully found categories
  - Red: No categories found
  - Blue: No application found
- Added yellow highlighting for headers
- Enhanced alert handling for invalid applications
- Renamed URL column to Application for clarity
- Updated status messaging for better error reporting
- Added cell color synchronization across rows

### Version 2.0.0 (Current)
- Added proxy and SSL certificate handling
- Implemented headless browser operation
- Enhanced Excel formatting with fixed column widths
- Added text wrapping and top alignment in Excel output
- Improved error handling with detailed error messages
- Added automatic directory creation
- Increased wait times for better reliability
- Added Chrome driver auto-installation

### Version 1.0.0 (Initial Release)
- Basic web scraping functionality
- Excel input/output support
- Category and subcategory extraction
- Simple error handling
- Basic Excel formatting

## 🚀 Features
- Automated web scraping from Palo Alto Networks Applipedia
- Batch processing of multiple URLs/applications
- Excel input/output with formatted results
- Headless browser operation
- Proxy and SSL certificate handling
- Auto-formatted Excel output with column width adjustment and text wrapping
- Robust error handling and retry mechanisms

## 🛠️ Prerequisites
```bash
python >= 3.7
Google Chrome browser
pip (Python package installer)
```

## 📦 Dependencies
```bash
selenium
pandas
openpyxl
webdriver_manager
typing
```

## ⚙️ Installation
1. Clone the repository:
```bash
git clone https://github.com/yourusername/applipedia-category-scraper.git
cd applipedia-category-scraper
```

2. Install requirements:
```bash
pip install -r requirements.txt
```

## 📁 Project Structure
```
applipedia-category-scraper/
├── Excel_Files/
│   ├── inputfile.xlsx  # Input URLs/applications
│   └── output.xlsx     # Scraped results
├── main.py            # Main script with scraping logic
└── README.md
```

## 📝 Input Format
Create `Excel_Files/inputfile.xlsx` with URLs/applications in the first column:
```
URL/Application
example.com
application-name
domain.com
```

## 🎮 Usage
1. Prepare input Excel file:
   - Place URLs or application names in the first column
   - Save as `Excel_Files/inputfile.xlsx`

2. Run the script:
```bash
python main.py
```

3. Find results in `Excel_Files/output.xlsx`

## 📊 Output Format
The script generates an Excel file with:
- SNO: Serial number
- URL: Input URL/application name
- Category: Extracted categories (multiple entries separated by newlines)
- Sub Category: Extracted subcategories (multiple entries separated by newlines)

## ⚠️ Error Handling
- Network issues: Configurable wait times and retry mechanisms
- Invalid URLs: Marked as "No Categories Found"
- SSL Certificate errors: Automatic handling with certificate acceptance
- Proxy issues: Configurable proxy settings with bypass options
- Missing directories: Automatic creation of Excel_Files directory

## 🔧 Configuration
Adjust these constants in `main.py`:
```python
CHROME_WAIT_TIME = 20  # Selenium wait time in seconds
BASE_URL = "https://applipedia.paloaltonetworks.com/"  # Base URL for scraping
```

## 📏 Excel Formatting
The output Excel file includes:
- Predefined column widths (A:5, B:20, C:35, D:35)
- Text wrapping enabled
- Top alignment for all cells
- Proper handling of multi-line content

## 🔒 Security Features
- Headless browser operation
- SSL certificate handling
- Proxy configuration options
- Sandbox disable options for Chrome
- Network timeout handling

## 🙏 Acknowledgments
- Palo Alto Networks Applipedia team
- Selenium WebDriver developers
- ChromeDriver development team
