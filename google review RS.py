from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import NoSuchElementException, ElementClickInterceptedException
import pandas as pd
import time

def scrape_google_maps_reviews(url, max_reviews=100):
    chrome_driver_path = "chromedriver-mac-x64/chromedriver"  # Ganti dengan path ke Chromedriver Anda
    service = Service(executable_path=chrome_driver_path)
    driver = webdriver.Chrome(service=service)

    driver.get(url)
    time.sleep(2)  # Tunggu beberapa detik hingga halaman terbuka

    reviews_data = []
    review_ids = set()

    try:
        # Scroll panel ulasan
        scrollable_div = driver.find_element(By.CSS_SELECTOR, ".m6QErb.DxyBCb.kA9KIf.dS8AEf")

        while len(reviews_data) < max_reviews:
            driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight", scrollable_div)
            time.sleep(2)

            # Temukan semua elemen ulasan
            reviews = driver.find_elements(By.XPATH, "//div[@data-review-id]")
            for review in reviews:
                if len(reviews_data) >= max_reviews:
                    break

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

                # Ekstrak data ulasan (ganti dengan CSS selector yang sesuai)
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

                # Ekstrak data tanggapan jika tersedia
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

        # Convert to DataFrame dan hapus duplikat
        df = pd.DataFrame(reviews_data, columns=['Reviewer Name', 'Review Date', 'Review Content', 'Review Rating', 'Responder Name', 'Response Date', 'Response Content']).drop_duplicates()
            
    except Exception as e:
        print(f"Error occurred: {e}")
        df = pd.DataFrame(columns=['Reviewer Name', 'Review Date', 'Review Content', 'Review Rating', 'Responder Name', 'Response Date', 'Response Content'])  # Create empty DataFrame

    finally:
        # Tutup browser
        driver.quit()

    return df

# URL lokasi Google Maps yang ingin Anda scrape ulasannya
url = "https://www.google.com/maps/place/Siloam+Hospitals+Kebon+Jeruk/@-6.1910578,106.7589568,17z/data=!4m8!3m7!1s0x2e69f71b743e853f:0xd437f33bf8bb0fca!8m2!3d-6.1910632!4d106.7638277!9m1!1b1!16s%2Fg%2F12156p2d?entry=ttu"

# Scrape ulasan dan simpan dalam DataFrame, batasi ke 500 ulasan
reviews_df = scrape_google_maps_reviews(url, max_reviews=100)

# Cetak DataFrame hasil
print(reviews_df)

# Simpan ke file Excel
reviews_df.to_excel("RS Siloam.xlsx", index=False)