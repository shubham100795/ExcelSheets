# ExcelSheets
Generate an excelsheet with random data and make modifications to the data generated

Setup :
Install python v3 to your system

run **python -V** from your command line to check pyhton version

Install the required module(openpyxl) using : **pip install openpyxl**

generate_random.py generates random data for students with marks in 3 subjects where user decides the number of records and the filename he wants to generate.

The generated excelsheet goes into excelsheets directory.

app.py shows all the excelsheets present inside excelsheets directory and is used to modify the marks of all the students to 70% of original value for any of the subjects present in the excelsheets.

A new column gets added to the excelsheet with the new values.

**A new sheet with name 'graph_sheet' gets generatedd in the same excel file which shows the bar graph of original and new values after modification.**

The modified excelsheet gets saved in modified_sheets directory.

Run :

**python generate_random.py** to generate a random record.

**python app.py** to modify marks in any of the 3 subjects.
