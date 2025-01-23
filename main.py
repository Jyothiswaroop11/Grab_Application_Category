import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from typing import List, Dict
import time
from datetime import datetime
import os
import openpyxl
from openpyxl.styles import Alignment

BASE_URL = "https://applipedia.paloaltonetworks.com/"
CHROME_WAIT_TIME = 20  # Increased wait time


class ExcelHandler:
    def read_excel(self, file_path: str) -> List[str]:
        try:
            df = pd.read_excel(file_path)
            # Get the first column that contains URLs
            urls_column = df.iloc[:, 0]
            return urls_column.dropna().tolist()
        except Exception as e:
            raise Exception(f"Error reading Excel: {str(e)}")

    def write_excel(self, data: List[Dict], output_file: str) -> None:
        try:
            df = pd.DataFrame(data)
            writer = pd.ExcelWriter(output_file, engine='openpyxl')

            df[['SNO', 'URL', 'Category', 'Sub Category']].to_excel(
                writer,
                index=False,
                sheet_name='Sheet1'
            )

            workbook = writer.book
            worksheet = writer.sheets['Sheet1']

            # # Auto-adjust columns width based on content
            # for column in worksheet.columns:
            #     max_length = 0
            #     column = [cell for cell in column]
            #     for cell in column:
            #         try:
            #             if len(str(cell.value)) > max_length:
            #                 max_length = len(str(cell.value))
            #         except:
            #             pass
            #     adjusted_width = (max_length + 2)
            #     worksheet.column_dimensions[column[0].column_letter].width = adjusted_width

            # Set column widths
            col_widths = {'A': 5, 'B': 20, 'C': 35, 'D': 35}
            for col, width in col_widths.items():
                worksheet.column_dimensions[col].width = width

            # Apply text wrapping and top alignment
            for row in worksheet.iter_rows(min_row=1, max_row=len(data) + 1):
                for cell in row:
                    cell.alignment = openpyxl.styles.Alignment(
                        vertical='top',
                        wrap_text=True
                    )

            writer.close()
        except Exception as e:
            raise Exception(f"Error writing Excel: {str(e)}")


class WebScraper:
    def __init__(self):
        self.driver = self._setup_driver()

    def _setup_driver(self) -> webdriver.Chrome:
        options = Options()
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--start-maximized")
        options.add_argument("--disable-gpu")
        options.add_argument("--disable-extensions")
        options.add_argument("--headless")

        # Use ChromeDriverManager to handle driver installation
        service = Service(ChromeDriverManager().install())

        driver = webdriver.Chrome(service=service, options=options)
        driver.set_page_load_timeout(CHROME_WAIT_TIME)
        return driver

    def search_and_extract(self, url: str) -> Dict:
        try:
            # Load website
            self.driver.get(BASE_URL)
            WebDriverWait(self.driver, CHROME_WAIT_TIME).until(
                EC.presence_of_element_located((By.TAG_NAME, "body"))
            )

            # Find and interact with search elements
            search_input = WebDriverWait(self.driver, CHROME_WAIT_TIME).until(
                EC.presence_of_element_located((By.ID, "tbSearch"))
            )
            search_button = WebDriverWait(self.driver, CHROME_WAIT_TIME).until(
                EC.element_to_be_clickable((By.ID, "btnSearch"))
            )

            # Perform search
            search_input.clear()
            search_input.send_keys(url)
            search_button.click()
            time.sleep(3)

            # Get all rows under CATEGORY and SUBCATEGORY sections
            categories = []
            subcategories = []

            # Find all category items
            category_elements = self.driver.find_elements(By.CSS_SELECTOR, "#CategoryList")
            for elem in category_elements:
                if elem.text.strip():
                    categories.append(elem.text.strip())

            # Find all subcategory items
            subcategory_elements = self.driver.find_elements(By.CSS_SELECTOR, "#SubCategoryList")
            for elem in subcategory_elements:
                if elem.text.strip():
                    subcategories.append(elem.text.strip())

            if categories and subcategories:
                return {
                    "URL": url,
                    "Category": "\n".join(categories),
                    "Sub Category": "\n".join(subcategories)
                }

            raise Exception("No results found")

        except Exception as e:
            print(f"Error processing URL {url}: {str(e)}")
            return {
                "URL": url,
                "Category": "No Categories Found",
                "Sub Category": "No Sub Categories Found"
            }

    def close(self):
        if self.driver:
            self.driver.quit()


def process_urls(input_file: str, output_file: str) -> None:
    excel_handler = ExcelHandler()
    web_scraper = WebScraper()

    try:
        print("Reading input file...")
        urls = excel_handler.read_excel(input_file)
        results = []

        print(f"Processing {len(urls)} URLs...")
        for index, url in enumerate(urls, 1):
            print(f"Processing URL {index}/{len(urls)}: {url}")
            result = web_scraper.search_and_extract(url)
            result['SNO'] = index
            results.append(result)
            time.sleep(2)  # Increased wait between requests

        print("Saving results...")
        excel_handler.write_excel(results, output_file)
        print(f"Results saved to {output_file}")

    except Exception as e:
        print(f"Error: {str(e)}")
    finally:
        web_scraper.close()


if __name__ == "__main__":
    input_file = "Excel_Files/inputfile.xlsx"
    output_file = "Excel_Files/output.xlsx"

    if not os.path.exists("Excel_Files"):
        os.makedirs("Excel_Files")

    process_urls(input_file, output_file)