# excel-filter-tool

📊 A Streamlit app to filter, group, and export Excel data with styled summaries.

---

## 📘 About the Project

This project was developed during my internship at **FERNUS**, an educational technology company based in Ankara, Turkey.

FERNUS needed a tool to quickly summarize Excel files related to book packaging. Manual analysis was slow and error-prone, so this app was built to automate the process through a simple interface.

It provides a clean interface to:
- Upload Excel files from the production workflow
- Group data by brand and category
- Classify books as **Katmanlı / Katmansız**
- Export styled summary reports in Excel format

Although it was originally built for internal use at FERNUS, the tool is general-purpose and can be adapted to other structured Excel workflows.

---

## 🚀 Features

- 📁 Upload Excel files (`.xlsx`) through a user-friendly web interface
- 📊 Automatically filter and group data by:
  - Brand (`MARKA ADI`)
  - Category (`KATEGORİ ADI`)
  - Packaging Status (`PAKETLEME DURUMU`)
- 🧠 Adds a computed label for PDF type: **Katmanlı / Katmansız**
- 📤 Export results into a single Excel file with:
  - Multiple sheets (`Kategori`, `Marka`, `Detaylı`)
  - Auto-sized columns and active filters
  - Colored headers and alternating row colors
  - Styled borders and centered text
- 🌐 Fully browser-based — no local Excel installation needed
- ⚙️ Built with only Python and open-source libraries

---

## 🛠 Tech Stack

- **[Streamlit](https://streamlit.io/)** – for building the interactive web interface
- **[pandas](https://pandas.pydata.org/)** – for data processing and grouping
- **[openpyxl](https://openpyxl.readthedocs.io/)** – for Excel export and styling
- **Python 3.10+**

---
