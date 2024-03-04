import selenium
from selenium import webdriver

import sys
import logging
import os, glob
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException
import time

# Configure logging
logging.basicConfig(filename='error.log', level=logging.ERROR)

def scrape_product_data(url):
    # Initialize the Chrome WebDriver with desired options
    chrome_options = webdriver.ChromeOptions()
    
    # Suppress certificate errors
    chrome_options.add_argument('--ignore-certificate-errors')
    
    # Set log level to suppress certificate-related messages
    chrome_options.add_argument('--log-level=3')
    
    # Initialize the Chrome WebDriver
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    
    driver.get(url)
    time.sleep(2)  # Add a delay to allow the page to load
    
    # Initialize an empty list to store product data
    product_data = []

    # Function to check if the forward arrow button is present and return its href attribute
    def get_next_page_href():
        try:
            forward_arrow_button = driver.find_element(By.XPATH, '//li[@class="arrow"]/a[text()="»"]')
            return forward_arrow_button.get_attribute('href')
        except NoSuchElementException:
            return None

    # Function to scrape product data
    def scrape_product_data_on_page():
        # Find all product elements
        product_elements = driver.find_elements(By.CLASS_NAME, 'product')

        # Loop through each product element
        for product_element in product_elements:
            # Extract product details
            product_title_element = product_element.find_element(By.CLASS_NAME, 'product-title')
            title = product_title_element.text
            link = product_title_element.get_attribute('href')

            # Extract the price
            product_price_element = product_element.find_element(By.CLASS_NAME, 'product-price')
            price_text = product_price_element.text
            prices = price_text.split('₹')
            original_price = prices[1] if len(prices) > 1 else None
            discounted_price = prices[2] if len(prices) > 2 else None

            # Check if the product has a review
            has_review = False
            try:
                review_element = product_element.find_element(By.CLASS_NAME, 'sr-only')
                if review_element.text == '5.0 star rating':
                    has_review = True
            except NoSuchElementException:
                pass
            
            has_image = False
            try:
                image_element = product_element.find_element(By.CSS_SELECTOR, '.product-thumb img')
                has_image = True
            except NoSuchElementException:
                pass

            # Append product data to the list
            product_data.append({
                'title': title,
                'link': link,
                'original_price': original_price,
                'discounted_price': discounted_price,
                'has_review': 'Yes' if has_review else 'No',
                'has_image': 'Yes' if has_image else 'No'
            })

    # Loop to navigate through all pages and scrape product data
    while True:
        # Scrape product data on the current page
        scrape_product_data_on_page()
        
        # Get the href of the next page
        next_page_href = get_next_page_href()
        if next_page_href:
            # Navigate to the next page
            driver.get(next_page_href)
            time.sleep(1)  # Add a delay to allow the page to load
        else:
            # If there is no next page, exit the loop
            break

    # Close the WebDriver
    driver.quit()

    return product_data

# Main function
def main():
    try:
        # Input URL
        url = input("Enter the URL of the website to scrape: ")

        # Scrape product data
        all_product_data = scrape_product_data(url)

        # Convert the scraped data to a DataFrame
        df = pd.DataFrame(all_product_data)
        excel_filename = os.path.basename(url)

        # Save the DataFrame to an Excel file with the extracted filename
        df.to_excel(f"{excel_filename}.xlsx", index=False)

        # Display output message
        print(f"Finished scraping. Scraped {len(df)} products.")
    except Exception as e:
        # Log any exceptions
        logging.error(f"An error occurred: {e}")
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    main()