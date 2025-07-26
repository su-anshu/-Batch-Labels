# 🏷️ Product Label Generator

A Streamlit web application for generating professional product labels from Excel files.

## Features

- **Excel File Upload**: Upload Excel files containing product names
- **Product Selection**: Choose products from a dropdown menu
- **Label Sizing**: Generate labels in two sizes:
  - 48x25mm (single label)
  - 96x25mm (two identical labels side by side)
- **Auto Date**: Current date automatically added in DD/MM/YYYY format
- **PDF Download**: Generate properly dimensioned PDF labels
- **Professional Layout**: Center-aligned text with bold product names

## Requirements

- Python 3.7+
- Streamlit
- Pandas
- ReportLab
- OpenPyXL

## Installation

1. Clone or download this repository
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

1. **Start the application**:
   ```bash
   streamlit run app.py
   ```
   Or double-click `run_app.bat` on Windows

2. **Prepare your Excel file**:
   - Ensure product names are in the first column
   - Supported formats: .xlsx, .xls

3. **Generate labels**:
   - Upload your Excel file
   - Select a product from the dropdown
   - Choose label size (48x25mm or 96x25mm)
   - Click "Generate Label PDF"
   - Download the generated PDF

## Excel File Format

Your Excel file should have product names in the first column:

| Product Name |
|-------------|
| Wireless Bluetooth Headphones |
| Smartphone Case - Black |
| USB-C Charging Cable |
| Portable Power Bank |

## Label Specifications

- **48x25mm**: Single label perfect for small products
- **96x25mm**: Two identical labels side by side for efficiency
- **Format**: Product name (large, bold) + Current date (DD/MM/YYYY)
- **Output**: High-quality PDF with exact dimensions

## Files

- `app.py`: Main Streamlit application
- `requirements.txt`: Python dependencies
- `run_app.bat`: Windows batch file to run the app
- `README.md`: This documentation

## Troubleshooting

- **File not loading**: Ensure your Excel file has data in the first column
- **PDF generation errors**: Check that product names don't contain special characters
- **Installation issues**: Make sure all dependencies are installed correctly

## License

This project is open source and available under the MIT License.