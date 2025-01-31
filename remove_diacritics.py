import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from openpyxl import load_workbook
import string
from unidecode import unidecode

# creating main window
window = tk.Tk()
window.title("Excel Diacritics Remover")

columns_var = tk.StringVar()

def uploadAction():
    try:
        # getting columns from input field
        columns = columns_var.get().strip().split(",")
        
        # crash if no columns selected
        if not columns or columns == [""]:
            print("Please enter column name(s) before uploading.")
            messagebox.showerror("Error", "Please enter column name(s) before uploading.")
            window.quit()
            return

        # clearing input tab
        columns_var.set("")

        # opening file dialog
        filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        
        # crash if no file selected
        if not filename:
            print("No file selected.")
            return

        # loading workbook and sheet
        wb = load_workbook(filename)

        # list name in excel
        ws = wb["List1"]

        # loading all columns
        all_columns = list(ws.columns)
        
        if all_columns:
            for cols in columns:
                # getting letter index in alphabet
                i = string.ascii_lowercase.index(cols.lower())
                for cell in all_columns[i]:
                    if cell.value:
                        cell.value = unidecode(cell.value)
        else:
            print("No data found in the sheet.")

        # changing filename based on .*
        if not filename.endswith("_modified.xlsx") and not filename.endswith("_modified.xls"):
            if filename.lower().endswith(".xlsx"):
                newFilename = filename.replace(".xlsx", "_modified.xlsx")
            elif filename.lower().endswith(".xls"):
                newFilename = filename.replace(".xls", "_modified.xls")
        else:
            newFilename = filename

        # saving new workbook
        wb.save(newFilename)

        # success
        print("modified file saved as: ", newFilename)
        messagebox.showinfo("Success", f"Modified file saved as:\n{newFilename}")
        window.destroy()

    except Exception as e:
        print("Error loading file:", e)
        messagebox.showerror("Error", f"An error occurred:\n{e}")
        window.destroy()

# GUI elements
columns_label = tk.Label(window, text="Columns:")
columns_label.grid(row=0, column=0, padx=10, pady=10)

columns_entry = tk.Entry(window, textvariable=columns_var)
columns_entry.grid(row=0, column=1, padx=10, pady=10)

button = tk.Button(window, text="Upload Excel File", command=uploadAction, width=25, height=2, bg="orange", fg="white")
button.grid(row=1, column=0, columnspan=2, pady=10)

# running the Tkinter event loop
window.mainloop()
