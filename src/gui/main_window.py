import tkinter as tk
from tkinter import filedialog


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Order Processor")

        self.excel_file_label = tk.Label(root, text="Excel File:")
        self.excel_file_label.pack()

        self.excel_file_entry = tk.Entry(root)
        self.excel_file_entry.pack()

        self.excel_file_button = tk.Button(root, text="Browse Excel", command=self.browse_excel)
        self.excel_file_button.pack()

        self.exclude_file_label = tk.Label(root, text="Exclude File:")
        self.exclude_file_label.pack()

        self.exclude_file_entry = tk.Entry(root)
        self.exclude_file_entry.pack()

        self.exclude_file_button = tk.Button(root, text="Browse Exclude", command=self.browse_exclude)
        self.exclude_file_button.pack()

        self.amount_label = tk.Label(root, text="Amount:")
        self.amount_label.pack()

        self.amount_entry = tk.Entry(root)
        self.amount_entry.pack()

        self.process_button = tk.Button(root, text="Process", command=self.process_files)
        self.process_button.pack()

    def browse_excel(self):
        excel_file = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        self.excel_file_entry.delete(0, tk.END)
        self.excel_file_entry.insert(0, excel_file)

    def browse_exclude(self):
        exclude_file = filedialog.askopenfilename(filetypes=[("Text Files", "*.txt")])
        self.exclude_file_entry.delete(0, tk.END)
        self.exclude_file_entry.insert(0, exclude_file)

    def process_files(self):
        excel_file = self.excel_file_entry.get()
        exclude_file = self.exclude_file_entry.get()
        amount = float(self.amount_entry.get())

        # Here you would call your processing logic using the provided files and amount
        # You'll need to implement the logic for reading and processing Excel files,
        # excluding certain names, calculating totals, and generating the output file.


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
