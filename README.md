Search automation with Selenium:

This is a Python code to search for product offers on two Brazilian e-commerce websites, using Selenium and Pandas libraries. The code searches for products on Google Shopping and Buscapé, filtering the results by a minimum and maximum price and a list of banned words in the product's name.

Libraries:

selenium is used to automate web browsers and interact with web pages. In this code, it is used to open the Chrome browser, navigate to the search pages and retrieve the results.

pandas is used for data manipulation and analysis. It is used to read the input file with the products to be searched, and store the results.

win32com.client is used to send an email with the results using Microsoft Outlook.

Input:

The input is a .xlsx file with a sheet named "produtos" and columns "Produto", "Termos Banidos", "Preço Mínimo" and "Preço Máximo". The first column contains the name of the products to be searched. The second column contains a list of words that should not be present in the product name. The third and fourth columns contain the minimum and maximum prices to filter the results.

Functions:

The code has two functions:

busca_google_shopping(nav, produto, termos_banidos, preco_minimo, preco_maximo)
This function searches for a product on Google Shopping, using the nav variable as the browser object, and the other parameters to configure the search. The function returns a list of tuples with the product name, price and link to the offer.

busca_buscape(nav, produto, termos_banidos, preco_minimo, preco_maximo)
This function searches for a product on Buscapé, using the nav variable as the browser object, and the other parameters to configure the search. The function returns a list of tuples with the product name, price and link to the offer.

Output:

The output is a .xlsx file with a sheet for each product searched, containing the results from Google Shopping and Buscapé. The sheet name is the same as the product name. Each sheet contains three columns: "Nome", "Preço" and "Link". The win32com.client library is used to send an email with the results to a recipient.

Usage:

To use this code, make sure you have the required libraries installed (selenium, pandas and win32com.client). Create an input file with the products to be searched and run the script. The output will be stored in a file named "output.xlsx" and an email will be sent with the results.

Note that the script uses the Chrome browser to perform the searches. Make sure you have the Chrome driver installed and set the path to the driver in the webdriver.Chrome() function.
