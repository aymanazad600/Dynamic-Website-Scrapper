# Dynamic Website Scraper

A Python-based web scraping tool that collects useful information from websites and exports the results to an Excel file.

The project was originally built to compare Kaggle datasets quickly, but it can be adapted for many other websites with only minor changes.

## Features

* Read URLs or search keywords from an input file
* Visit each webpage automatically
* Extract useful page information
* Export everything into a structured Excel spreadsheet
* Continue running even if one page fails
* Works on both macOS and Windows

## Project Structure

```
Dynamic-Website-Scrapper/
│
├── agent.py
├── requirements.txt
├── urls.txt
├── output/
├── Screenshots/
└── README.md
```

## Installation

Clone the repository.

Create a virtual environment.

**macOS**

```bash
python3 -m venv .venv
source .venv/bin/activate
```

**Windows**

```bash
python -m venv .venv
.venv\Scripts\activate
```

Install the required packages.

```bash
pip install -r requirements.txt
```

## How to Use

1. Open `urls.txt`.
2. Add one website URL (or multiple URLs, one per line).
3. Save the file.
4. Run:

```bash
python agent.py
```

When the scraping is finished, the Excel file will be generated inside the `output` folder.

## Notes

* The scraper is designed for publicly accessible pages.
* Some websites use anti-bot protection or require authentication. In those cases, additional handling may be required.
* The scraper is modular, making it easy to customize for different websites.

## Example Output

The `Screenshots` folder contains example search pages and the generated Excel output.

## Future Improvements

* Search using keywords instead of URLs
* Automatic website structure detection
* Support for additional website layouts
* More export formats (CSV, JSON)

