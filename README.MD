# ExcelHandler

ExcelHandler is a Python class that provides a flexible way to read and update Excel documents using either the openpyxl or xlwings library, allowing you to choose the library that best suits your needs.

## Usage

### Initialization

Create an instance of the ExcelHandler class by specifying the Excel file path and the library to use (either 'openpyxl' or 'xlwings'):

```python
from excel_handler import ExcelHandler

# Specify the Excel file path and library choice
excel_handler = ExcelHandler("example.xlsx", library='openpyxl')
```

## Reading Excel Data

To read data from an Excel sheet, use the `read_excel` method:

```python
sheet_name = "Sheet1"
data = excel_handler.read_excel(sheet_name)
print(data)
```

## Updating Excel Data

To update data in an Excel sheet, use the `update_excel` method:

```python
sheet_name = "Sheet1"
new_data = [["Updated", "Data"], [1, 2], [3, 4]]
excel_handler.update_excel(sheet_name, new_data)
```


## Requirements

This project has the following requirements:

- Python 3.x
- `openpyxl` library (for openpyxl support)
- `xlwings` library (for xlwings support)

To use this project, make sure you have Python 3.x installed on your system. You can download Python from the official website: [Python Downloads](https://www.python.org/downloads/).

Additionally, you'll need to install the required libraries using pip. You can do this by running the following commands:

```bash
pip install openpyxl xlwings
```
These libraries are necessary for working with Excel files using the ExcelHandler class.

## Installation

To get started with this project, follow the steps below:

1. Clone the repository to your local machine:

```bash
git clone https://github.com/yourusername/excel-handler.git
cd excel-handler
```

Install the required libraries as mentioned in the "Requirements" section above:

```bash
pip install openpyxl xlwings
```
