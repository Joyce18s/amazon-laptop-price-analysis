from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import re
import unicodedata
import pandas as pd

# opening site
driver = webdriver.Chrome()
driver.get("https://www.amazon.in")

wait = WebDriverWait(driver, 15)

# searching "laptop" in searchbar
search = wait.until(
    EC.presence_of_element_located((By.ID, "twotabsearchtextbox"))
)
search.send_keys("laptop")

search_button = wait.until(
    EC.element_to_be_clickable((By.ID, "nav-search-submit-button"))
)
search_button.click()

# loop to scrape all product name and price
seen_titles = set()
data = []

while True:
    time.sleep(3)

    products = wait.until(
        EC.presence_of_all_elements_located(
            (By.CSS_SELECTOR, "div.s-result-item[data-component-type='s-search-result']")
        )
    )

    for p in products:
        try:
            name = p.find_element(By.CSS_SELECTOR, "h2 span").text

            if name in seen_titles:
                continue

            seen_titles.add(name)

            full_text = p.text
            match = re.search(r'₹[\d,]+', full_text)

            if match:
                price = match.group()
                price = price.replace("₹", "").replace(",", "")
                price = int(price)
            else:
                price = "N/A"

            data.append([name, price])

        except:
            continue

    try:
        next_button = wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "a.s-pagination-next"))
        )
        next_button.click()
        time.sleep(2)

    except:
        print("No more pages")
        break


#converting data into Dataframe
df = pd.DataFrame(data, columns=["PRODUCT NAME", "PRICE"])

#  Extracting brand name an adding "Brand" column in ataframe
import re
import unicodedata

def get_brand(title):
    title = str(title)

    title = unicodedata.normalize("NFKD", title)
    title = title.lower()

    if "thinkpad" in title:
        return "Lenovo"

    if "h.p" in title or " hp " in title:
        return "HP"
    if "portege" in title or "portégé" in title:
        return "Dynabook"

    brand_map = {
        "dell": "Dell",
        "hp": "HP",
        "alienware": "ALIENWARE",
        "lenovo": "Lenovo",
        "asus": "Asus",
        "acer": "Acer",
        "apple": "Apple",
        "msi": "MSI",
        "samsung": "Samsung",
        "primebook": "Primebook",
        "honor": "Honor",
        "avita": "Avita",
        "chuwi": "Chuwi",
        "lg": "LG"
    }

    for key, value in brand_map.items():
        if key in title:
            return value

    return "Unknown"

df["BRAND"] = df["PRODUCT NAME"].apply(get_brand)   

# Saving datafram in excel
df.to_excel("Laptop_Data.xlsx", index = False)
print(df)