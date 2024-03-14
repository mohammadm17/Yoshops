# Yoshops Web Scraper

This is a web scraping tool built using Selenium that extracts product information from the Yoshops website. It scrapes details such as product name, price, discounted price, link, and reviews, and stores the data in an Excel sheet for further analysis.

## Installation

1. Clone or download this repository to your local machine.
2. Ensure you have Python installed on your system.
3. Install the required dependencies using pip:


4. Ensure you have the Chrome WebDriver installed. You can download it from [here](https://chromedriver.chromium.org/downloads) and place it in the project directory.

## Usage

To scrape a specific category from Yoshops:

1. Run the `yoshops_scraper.py` script.
2. Enter the URL of the category you want to scrape when prompted.
3. Wait for the scraper to collect the data. It will be stored in an Excel file named `yoshops_data.xlsx`.

## Example

Here's an example of how to use the scraper:

python yoshops_scraper.py
Enter the URL of the category you want to scrape: https://yoshops.com/t/toys

## Contributing

If you'd like to contribute to this project, feel free to fork the repository and submit a pull request with your changes.

## License

This project is licensed under the [MIT License](LICENSE).
