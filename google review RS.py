from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import NoSuchElementException, ElementClickInterceptedException
from selenium.webdriver.common.keys import Keys
import pandas as pd
import openpyxl
import time

# ChromeDriver path
chrome_driver_path = "chromedriver-mac-x64/chromedriver"

# Setting up ChromeDriver
service = Service(executable_path=chrome_driver_path)
driver = webdriver.Chrome(service=service)

def scrape_google_maps_reviews(url):
    driver.get(url)
    time.sleep(2)  # Wait for the page to load

    reviews_data = []
    review_ids = set()  # Keep track of review IDs to avoid duplicates

    try:
        # Scroll the reviews panel
        scrollable_div = driver.find_element(By.CSS_SELECTOR, ".m6QErb.DxyBCb.kA9KIf.dS8AEf")
        last_height = driver.execute_script("return arguments[0].scrollHeight", scrollable_div)
        
        while True:
            driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight", scrollable_div)
            time.sleep(2)
            new_height = driver.execute_script("return arguments[0].scrollHeight", scrollable_div)
            if new_height == last_height:
                break
            last_height = new_height

        # Find all review elements
        reviews = driver.find_elements(By.XPATH, "//div[@data-review-id]")
        for review in reviews:
            review_id = review.get_attribute("data-review-id")
            if review_id in review_ids:
                continue
            review_ids.add(review_id)

            try:
                see_more_button = review.find_element(By.CSS_SELECTOR, "button.w8nwRe.kyuRq")
                see_more_button.click()
                time.sleep(0.5)
            except (NoSuchElementException, ElementClickInterceptedException):
                pass

            # Extract data
            try:
                reviewer_name = review.find_element(By.CSS_SELECTOR, "div.d4r55").text
            except NoSuchElementException:
                reviewer_name = "unknown name"
            try:
                review_date = review.find_element(By.CSS_SELECTOR, "span.rsqaWe").text
            except NoSuchElementException:
                review_date = "unknown date"
            try:
                review_content = review.find_element(By.CSS_SELECTOR, "span.wiI7pd").text
            except NoSuchElementException:
                review_content = "unknown content"
            try:
                review_rating = review.find_element(By.CSS_SELECTOR, "span.kvMYJc").get_attribute("aria-label")
            except NoSuchElementException:
                review_rating = "unknown rating"

            # Extract response data if available
            try:
                response = review.find_element(By.CSS_SELECTOR, "div.CDe7pd")
                responder_name = response.find_element(By.CSS_SELECTOR, "span.nM6d2c").text if response.find_element(By.CSS_SELECTOR, "span.nM6d2c") else "No response"
                response_date = response.find_element(By.CSS_SELECTOR, "span.DZSIDd").text if response.find_element(By.CSS_SELECTOR, "span.DZSIDd") else "No response date"
                response_content = response.find_element(By.CSS_SELECTOR, "div.wiI7pd").text if response.find_element(By.CSS_SELECTOR, "div.wiI7pd") else "No response content"
            except NoSuchElementException:
                responder_name = "No response"
                response_date = "No response date"
                response_content = "No response content"

            reviews_data.append([reviewer_name, review_date, review_content, review_rating, responder_name, response_date, response_content])

        # Convert to DataFrame and remove duplicates
        df = pd.DataFrame(reviews_data, columns=['Reviewer Name', 'Review Date', 'Review Content', 'Review Rating', 'Responder Name', 'Response Date', 'Response Content']).drop_duplicates()
            
    except Exception as e:
        print(f"Error occurred: {e}")
        df = pd.DataFrame(columns=['Reviewer Name', 'Review Date', 'Review Content', 'Review Rating', 'Responder Name', 'Response Date', 'Response Content'])  # Create empty DataFrame

    # Close the browser
    driver.quit()

    return df

# URL of the Google Maps location
url = "https://www.google.com/maps/place/RSIA+Plamongan+Indah/@-7.0235972,110.4952163,17z/data=!4m16!1m7!3m6!1s0x2e708d97e40bbb15:0x56675269b5402a31!2sRSIA+Plamongan+Indah!8m2!3d-7.0236841!4d110.4979331!16s%2Fg%2F11b6_dn47d!3m7!1s0x2e708d97e40bbb15:0x56675269b5402a31!8m2!3d-7.0236841!4d110.4979331!9m1!1b1!16s%2Fg%2F11b6_dn47d?entry=ttu"

# Scrape the reviews and store in a DataFrame
reviews_df = scrape_google_maps_reviews(url)

# Print the resulting DataFrame
print(reviews_df)


reviews_df.to_excel("WebScrape.xlsx", index=False)