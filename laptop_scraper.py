import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook

# Constants
BASE_URL = "https://webscraper.io/test-sites/e-commerce/static/computers/laptops"
USD_TO_NGN = 1500
MAX_ENTRIES = 500

def scrape_to_excel(output_file="webscraper_products.xlsx"):
    all_products = []
    page = 1

    while len(all_products) < MAX_ENTRIES:
        url = f"{BASE_URL}?page={page}" if page > 1 else BASE_URL
        print(f"Scraping page {page}...")

        response = requests.get(url)
        if response.status_code != 200:
            print(f"Failed to load page {page}")
            break

        soup = BeautifulSoup(response.text, "html.parser")
        items = soup.select(".thumbnail")
        if not items:
            break  # No more data

        for item in items:
            if len(all_products) >= MAX_ENTRIES:
                break

            title = item.select_one(".title").text.strip()
            price_text = item.select_one(".price").text.strip().replace("$", "")
            try:
                price_usd = float(price_text)
                price_ngn = round(price_usd * USD_TO_NGN, 2)
            except:
                price_usd = 0.0
                price_ngn = 0.0
            description = item.select_one(".description").text.strip()

            all_products.append([title, price_usd, price_ngn, description])

        page += 1

    # Write to .xlsx using openpyxl
    wb = Workbook()
    ws = wb.active
    ws.title = "Laptops"

    # Write header
    ws.append(["Title", "Price (USD)", "Price (NGN)", "Description"])

    # Write data rows
    for product in all_products:
        ws.append(product)

    wb.save(output_file)
    print(f"\n✅ Scraped {len(all_products)} products")
    print(f"📄 Excel file saved to: {output_file}")


if __name__ == "__main__":
    scrape_to_excel()
