import win32com.client as win32
import tkinter as tk
from tkinter import messagebox, Listbox
import webbrowser
from decimal import Decimal, getcontext
import math

def format_selected_cells():
    try:
        excel = win32.GetActiveObject("Excel.Application")
        selection = excel.Selection
        if isinstance(selection, win32.CDispatch):
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
        else:
            messagebox.showwarning("Selection Error", "Please select one or more cells.")

    except Exception as e:
        messagebox.showerror("Error", str(e))

def search_wolfram_alpha():
    try:
        excel = win32.GetActiveObject("Excel.Application")
        selection = excel.Selection
        if isinstance(selection, win32.CDispatch) and selection.Count == 1:
            cell_value = selection.Text
            url = f"http://www.wolframalpha.com/input/?i={cell_value}"
            webbrowser.open_new_tab(url)
            messagebox.showinfo("Wolfram Alpha Search", f"Searching for: {cell_value}")
        else:
            messagebox.showwarning("Selection Error", "Please select a single cell.")

    except Exception as e:
        messagebox.showerror("Error", str(e))

def update_list():
    try:
        excel = win32.GetActiveObject("Excel.Application")
        selection = excel.Selection
        if isinstance(selection, win32.CDispatch):
            listbox.delete(0, tk.END)
            for cell in selection:
                listbox.insert(tk.END, cell.Text)
        else:
            messagebox.showwarning("Selection Error", "Please select one or more cells.")

    except Exception as e:
        messagebox.showerror("Error", str(e))

root = tk.Tk()
root.title("Excel Tools")
root.geometry("400x600")

listbox = Listbox(root)
listbox.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

update_btn = tk.Button(root, text="Update List", command=update_list)
update_btn.pack(padx=10, pady=5)

format_btn = tk.Button(root, text="Format to Scientific Notation", command=format_selected_cells)
format_btn.pack(padx=10, pady=5)

search_btn = tk.Button(root, text="Search in Wolfram Alpha", command=search_wolfram_alpha)
search_btn.pack(padx=10, pady=5)

root.mainloop()