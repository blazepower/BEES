# BEES

This repository contains two scripts developed for the BEES EFRC
All scripts require Python to be installed

---

## Converter
Converter takes the BEES Active Publications List Excel Spreadsheet and outputs a Word Document titled `Appendix.docx`.

A couple notes on this:
* Ensure that the Active Publications List is up to date and in the same folder as the python script
* Ensure that there is not another file named `Appendix.docx`
* Python library to access Excel spreadsheet may be deprecated and an older version may need to be used to get full functionality

---
## Generator
Generator takes the BEES Publication Database and outputs HTML in the same format as needed by the BEES Website

A couple notes on this:
* Ensure that the BEES Publications Database Excel spreadsheet is up to date and in the same folder as the python script
* Modify row bounds on line 8 to only output HTML for rows that are needed
* Ensure openpyxl library is installed e.g.: `pip install openpyxl`
