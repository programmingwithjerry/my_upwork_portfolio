import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference

# ----------------------------- CONFIGURATION -----------------------------

# Base URL of the target website
BASE_URL = "https://webscraper.io/test-sites/e-commerce/static/computers/laptops"

# USD to NGN exchange rate (you can update this as needed)
USD_TO_NGN = 1500

# Maximum number of products to scrape
MAX_ENTRIES = 500

# Output Excel file
OUTPUT_FILE = "webscraper_products_with_chart.xlsx"

# ------------------------------ SCRAPER FUNCTION ------------------------------

def scrape_to_excel_with_chart(output_file: str = OUTPUT_FILE):
    all_products = []
    page = 1

    # Loop through paginated product pages until we hit the limit or run out of data
    while len(all_products) < MAX_ENTRIES:
        url = f"{BASE_URL}?page={page}" if page > 1 else BASE_URL
        print(f"Scraping page {page}...")

        # Send HTTP GET request
        response = requests.get(url)
        if response.status_code != 200:
            print(f"Failed to load page {page}")
            break

        # Parse the page content
        soup = BeautifulSoup(response.text, "html.parser")
        items = soup.select(".thumbnail")
        if not items:
            break  # No more products

        # Extract title, price, description
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

    # ------------------------------ EXCEL WRITING ------------------------------

    # Create a new Excel workbook and worksheet
    wb = Workbook()
    ws = wb.active
    ws.title = "Laptops"

    # Header row
    ws.append(["Title", "Price (USD)", "Price (NGN)", "Description"])

    # Write each row of product data
    for product in all_products:
        ws.append(product)

    # ------------------------------ CHART GENERATION ------------------------------

    # Create a bar chart of the top 10 most expensive products
    chart = BarChart()
    chart.title = "Top 10 Laptop Prices (USD)"
    chart.x_axis.title = "Laptop"
    chart.y_axis.title = "Price (USD)"
    chart.width = 20
    chart.height = 10

    # Select top 10 rows for chart
    top_rows = min(10, len(all_products))

    # Reference: (values=Y-axis, categories=X-axis)
    values = Reference(ws, min_col=2, min_row=2, max_row=top_rows + 1)
    labels = Reference(ws, min_col=1, min_row=2, max_row=top_rows + 1)

    chart.add_data(values, titles_from_data=False)
    chart.set_categories(labels)

    # Insert the chart after the data
    chart_position = f"E2"
    ws.add_chart(chart, chart_position)

    # Save the workbook
    wb.save(output_file)

    print(f"\n✅ Scraped {len(all_products)} products")
    print(f"📊 Excel with chart saved as: {output_file}")

# ------------------------------ RUN ------------------------------

if __name__ == "__main__":
    scrape_to_excel_with_chart()
