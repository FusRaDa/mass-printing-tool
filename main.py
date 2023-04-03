import tkinter as tk
from tkinter import messagebox
from openpyxl import Workbook


def generate_excel_file():
    warning_title = "File Execution Failed"

    if location.get().__contains__(" ") or len(location.get()) == 0 \
            or item_number.get().__contains__(" ") or len(item_number.get()) == 0 \
            or lot.get().__contains__(" ") or len(lot.get()) == 0 \
            or quantity.get().__contains__(" ") or len(quantity.get()) == 0 \
            or number_of_labels.get().__contains__(" ") or len(number_of_labels.get()) == 0:
        messagebox.showerror(warning_title, "Fill all inputs, values must not contain spaces.")
        return

    print('executed')






# run ui
window = tk.Tk()
window.title("Mass Printing Generator")

tk.Label(window, text="Location").grid(row=0, column=1)
tk.Label(window, text="Item #").grid(row=0, column=2)
tk.Label(window, text="Lot").grid(row=0, column=3)
tk.Label(window, text="Quantity").grid(row=0, column=4)

tk.Label(window, text="Number of Labels").grid(row=2, column=2, pady=(15, 1))

tk.Button(window, text="Generate Excel File", command=generate_excel_file).grid(row=3, column=3)

location = tk.Entry(window)
item_number = tk.Entry(window)
lot = tk.Entry(window)
quantity = tk.Entry(window)

number_of_labels = tk.Entry(window)

location.grid(row=1, column=1)
item_number.grid(row=1, column=2)
lot.grid(row=1, column=3)
quantity.grid(row=1, column=4)

number_of_labels.grid(row=3, column=2)

window.mainloop()
