# NISApp

NISApp is a desktop application for managing and processing aircraft part and maintenance data. It provides a graphical user interface (GUI) for reading, processing, and exporting data from CSV files, and generating formatted Word documents. The application is built with Python and uses libraries such as Tkinter, pandas, python-docx, and others.

## Features
- Load and process aircraft part data from CSV files
- Search and filter part numbers and descriptions
- Calculate and aggregate time/cycle data for maintenance records
- Export formatted reports to Microsoft Word (.docx)
- User-friendly GUI with Tkinter and PySimpleGUI
- Support for Turkish Airlines data formats

## Requirements
- Python 3.7+
- Required Python packages:
  - pandas
  - python-docx
  - PySimpleGUI
  - Pillow
  - tkcalendar
  - babel
  - PyPDF2

## Installation
1. Clone the repository:
   ```sh
   git clone https://github.com/enesdmrkl/NISApp-repo.git
   ```
2. Install the required packages:
   ```sh
   pip install -r requirements.txt
   ```
   Or install manually:
   ```sh
   pip install pandas python-docx PySimpleGUI Pillow tkcalendar babel PyPDF2
   ```

## Usage
1. Place your data files (`sample.csv`, `msn2.csv`) in the project directory.
2. Run the application:
   ```sh
   python NISApp.py
   ```
3. Use the GUI to load, process, and export your data.

## File Structure
- `NISApp.py` - Main application script
- `sample.csv` - Sample part data
- `msn2.csv` - Aircraft serial number data
- `icon.ico`, `turkishairlines.png`, `word.docx` - Assets and templates

## License
This project is licensed under the MIT License.

## Author
- [enesdmrkl](https://github.com/enesdmrkl)
