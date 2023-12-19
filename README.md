<h1>Automated Product Registration</h1>

> Here is a Python script that automates the product registration process in an Excel spreadsheet. The script uses the openpyxl library to access and manipulate the spreadsheet, and the pyperclip and pyautogui libraries to copy and paste information between the spreadsheet and an external application.

Prerequisites

Make sure to have the following libraries installed before running the script:

```bash
pip install openpyxl pyperclip pyautogui
```

How to use


1- Open the Excel spreadsheet 'produtos_ficticios.xlsx' in the same directory as the script.
2- Execute the Python script.

The script simulates data entry into specific fields of an external application for each product listed in the spreadsheet. It will follow these steps for each product:

1- Product Name: Copy and paste the product name.
2- Description: Copy and paste the description.
3- Category: Copy and paste the category.
4- Product Code: Copy and paste the product code.
5- Weight: Copy and paste the weight.
6- Dimensions: Copy and paste the dimensions.
7- Price: Copy and paste the price.
8- Stock Quantity: Copy and paste the stock quantity.
9- Expiry Date: Copy and paste the expiry date.
10- Color: Copy and paste the color.
11- Size: Select the size as specified.
12- Material: Copy and paste the material.
13- Manufacturer: Copy and paste the manufacturer.
14- Country of Origin: Copy and paste the country of origin.
15- Observations: Copy and paste observations.
16- Barcode: Copy and paste the barcode.
17- Warehouse Location: Copy and paste the warehouse location.
18- Conclusion and Confirmation: Click on buttons to conclude and confirm the product inclusion.

The script repeats these steps for each product in the spreadsheet, advancing to the next page when necessary. After completing all products, the process is finished.

Note: This script is designed for a specific application and may require adjustments to work in different environments. Make sure to understand the operation of the target application before running the script. Use it responsibly and in accordance with the terms of service of the application.
