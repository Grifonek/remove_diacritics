# Excel Diacritics Remover

This is a simple Tkinter-based Python application that removes diacritics from text in a selected Excel file. The modified file is saved with \_modified appended to its name.

## 📌 Features

✅ Select specific columns to process\
✅ Removes diacritics using unidecode\
✅ Supports .xlsx and .xls Excel files\
✅ User-friendly Tkinter GUI\
✅ Displays success or error messages\
✅ Automatically saves a new file

## 📂 Project Structure

📁 Excel-Diacritics-Remover\
│── remove_diacritics.py # Main Python script\
│── requirements.txt # Required dependencies\
└── README.md # Project documentation

## 🛠 Installation & Setup

1️⃣ Install Python (if not already installed)

    Download and install Python from python.org
    Make sure to check ✅ "Add Python to PATH" during installation

2️⃣ Install Dependencies

Run the following command in the terminal (inside the project folder):

    pip install -r requirements.txt

3️⃣ Run the Application

## Usage

You can run the script this way:

✔ Option 1: Using Python

    python remove_diacritics.py

## 📖 Usage Guide

1️⃣ Enter the column letters (e.g., A,C,D,a,c,d) in the input field.\
2️⃣ Click "Upload Excel File" and select your .xlsx or .xls file.\
3️⃣ The script will process the selected columns and remove diacritics.\
4️⃣ A new file with \_modified in its name will be saved in the same folder.\
5️⃣ A success message will confirm the operation.

## ⚙ Requirements

Python 3.x\
openpyxl (for handling Excel files)\
unidecode (for removing diacritics)

If missing, install them using:

    pip install openpyxl unidecode

## 🛑 Troubleshooting

❌ Error: "No module named openpyxl/unidecode"\
✅ Run pip install -r requirements.txt

❌ File is open in another program\
✅ Close the Excel file before running the script.

❌ Columns input is empty\
✅ Enter at least one column before uploading the file.

## 📜 License

This project is open-source and free to use.
