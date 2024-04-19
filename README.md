# Excel Sorting and Matching Script

This script reads an Excel file, sorts the data by site name, site address, and latlong, and creates three separate sheets in a new Excel file. It also finds exact duplicates and high matches based on the site name and prints the matching attributes.

## Prerequisites

- Python 3.x
- `pandas` library
- `fuzzywuzzy` library

## Setup

1. Create a virtual environment (optional but recommended): `python -m venv myenv`

2. Activate the virtual environment:
- For Windows: `myenv\Scripts\activate`
- For macOS and Linux: `source myenv/bin/activate`

3. Install the required libraries: `pip install -r requirements.txt`

## Usage

1. Place your input Excel file in the same directory as the script.

2. Open the script in a text editor and replace `'input.xlsx'` with the actual filename of your input Excel file.

3. Run the script: `python app.py`

4. The script will generate an output Excel file named `'output.xlsx'` with the sorted data in three separate sheets.

5. The script will also print the exact duplicates and high matches along with the matching attributes.

## Deactivating the Virtual Environment

When you're done running the script, you can deactivate the virtual environment by running the following command: