# excel-filter-tool

A lightweight Excel filtering tool for tracking book packaging progress in publishing workflows.

---

## About the Project

This project was developed during my internship at **FERNUS**, an educational technology company based in Ankara, Turkey.

FERNUS needed a tool to track the packaging status of books by brand and category using data stored in Excel files.  
This tool was built to automate that process and provide a clear overview of which books have been processed and which ones are still pending.

It was specifically designed for internal use at FERNUS to simplify the tracking of packaging workflows across different brands and categories.

The tool offers a simple web-based interface that allows users to:
- Upload Excel files from the book production pipeline
- Group and filter data by brand, category, and packaging status
- Classify books as either "Katmanlı" or "Katmansız"
- Export the results as a well-formatted Excel report

---

## Features

- Upload Excel files (`.xlsx`) via a user-friendly web interface
- Automatically group and filter data by:
  - Brand (`MARKA ADI`)
  - Category (`KATEGORİ ADI`)
  - Packaging Status (`PAKETLEME DURUMU`)
- Automatically assign a "Katmanlı" or "Katmansız" label based on PDF metadata
- Export the output to a single Excel file that includes:
  - Multiple sheets (`Kategori`, `Marka`, `Detaylı`)
  - Styled headers, alternating row colors, borders, and filters
  - Auto-adjusted column widths
- Runs entirely in the browser—no Excel installation required
- Built using Python and open-source libraries only

---

## Tech Stack

- [Streamlit](https://streamlit.io/) – interactive web interface
- [pandas](https://pandas.pydata.org/) – data manipulation and aggregation
- [openpyxl](https://openpyxl.readthedocs.io/) – Excel export and styling
