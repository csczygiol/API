import openpyxl
import requests
from bs4 import BeautifulSoup

url = "https://www.tagesschau.de/wirtschaft/boersenkurse/dax-index-846900/"

# GET Request
response = requests.get(url)

if response.status_code == 200:
    # Parse content
    soup = BeautifulSoup(response.text, 'html.parser')

    # Find the specified elements
    price_element = soup.find('span', class_='price')
    price_currency_element = soup.find('span', class_='priceCurrency')
    change_element = soup.find('span', class_='change')

    # Extract text content
    price = price_element.text.strip()
    price_currency = price_currency_element.text.strip()
    change = change_element.text.strip()
    
    # Load Excelsheet
    workbook = openpyxl.load_workbook("apicalldax.xlsx")
    sheet = workbook.active
    
    # Create new row to append the data
    new_data = [price, price_currency, change]
    
    # Append new data to the next empty row
    sheet.append(new_data)

    workbook.save("apicalldax.xlsx")
    
    print("Data has been stored in apicalldax.xlsx")
else:
    print("Failed to retrieve the web page. Status code:", response.status_code)
