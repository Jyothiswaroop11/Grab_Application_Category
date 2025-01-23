# 🎯 Grab_Application_Category

## 📋 Description
A web scraping tool to automatically extract category and subcategory information from Palo Alto Networks Applipedia for given applications.

## 🚀 Features
- Automated web scraping from Applipedia
- Batch processing of multiple applications
- Excel input/output support
- Headless browser operation
- Auto-formatted Excel output with proper alignment and text wrapping

## 🛠️ Prerequisites
```bash
python >= 3.7
pip
Chrome browser
```

## 📦 Dependencies
```bash
selenium
pandas
openpyxl
webdriver_manager
```

## ⚙️ Installation
1. Clone the repository:
```bash
git clone https://github.com/yourusername/Grab_Application_Category.git
cd Grab_Application_Category
```

2. Install requirements:
```bash
pip install -r requirements.txt
```

## 📁 Project Structure
```
Grab_Application_Category/
├── Excel_Files/
│   ├── inputfile.xlsx
│   └── output.xlsx
├── main.py
└── README.md
```

## 📝 Input Format
Create `Excel_Files/inputfile.xlsx` with applications in first column:
```
Application
adobe
abs
cisco
```

## 🎮 Usage
1. Prepare input Excel file:
   - Place URLs in first column
   - Save as `Excel_Files/inputfile.xlsx`

2. Run script:
```bash
python main.py
```

3. Find results in `Excel_Files/output.xlsx`

## 📊 Output Format
The script generates an Excel file with:
- SNO: Serial number
- URL: Application name
- Category: Application categories
- Sub Category: Application subcategories

## ⚠️ Error Handling
- Invalid URLs: Marked as "Not Found"
- Network issues: Error messages in console
- Missing input file: Directory creation prompt

## 🔧 Customization
Adjust these constants in `main.py`:
```python
CHROME_WAIT_TIME = 20  # Selenium wait time
BASE_URL = "https://applipedia.paloaltonetworks.com/"
```

## 🙏 Acknowledgments
- Palo Alto Networks for Applipedia
- Selenium WebDriver team
- ChromeDriver developers
