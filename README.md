# moreover_scraper
Web scraper for Moreover

## Dependancies
* Selenium
* Google chrome (requires desktop env even when using in headless mode)
* Openpyxl

The scraper reads a list of 21236 urls from a xlsx file and automatically searches the Moreover site for the url.
Title and language are grabbed and inserted into the xlsx file, then saved as an updated version.
