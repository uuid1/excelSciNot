import win32com.client as win32
import tkinter as tk
from tkinter import messagebox
import webbrowser
from decimal import Decimal, getcontext
import math

def format_selected_cells():
    try:
        excel = win32.GetActiveObject("Excel.Application")
        workbook = excel.ActiveWorkbook
        sheet = workbook.ActiveSheet
        selection = excel.Selection

        getcontext().prec = 30  # Set decimal precision

        for cell in selection:
            if cell.Value is not None and isinstance(cell.Value, (float, int)):
                value = Decimal(str(cell.Value))
                if value != 0:
                    power = math.floor(math.log10(abs(value)))
                    mantissa = value / (Decimal(10) ** power)
                    formatted_value = "{:.2f} * 10^{}".format(mantissa, power)
                    cell.Value = formatted_value
        messagebox.showinfo("Success", "Selected cells have been formatted to scientific notation.")
    except Exception as e:
        messagebox.showerror("Error", str(e))

def search_wolfram_alpha():
    try:
        excel = win32.GetActiveObject("Excel.Application")
        selection = excel.Selection

        if selection.Count == 1:
            cell_value = selection.Value
            search_query = str(cell_value)
            url = f"https://www.wolframalpha.com/input/?i={search_query}"
            webbrowser.open_new_tab(url)
            messagebox.showinfo("Wolfram Alpha Search", f"Searching for: {search_query}")
        else:
            messagebox.showwarning("Selection Error", "Please select a single cell.")
    except Exception as e:
        messagebox.showerror("Error", str(e))

def main():
    root = tk.Tk()
    root.title("Excel Tool")

    format_btn = tk.Button(root, text="Format to Scientific Notation", command=format_selected_cells)
    format_btn.pack(padx=10, pady=5)

    search_btn = tk.Button(root, text="Search in Wolfram Alpha", command=search_wolfram_alpha)
    search_btn.pack(padx=10, pady=5)

    root.mainloop()

if __name__ == "__main__":
    main()