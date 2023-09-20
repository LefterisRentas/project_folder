import datetime
from tkinter import filedialog
import tkinter as tk
from tkcalendar import Calendar

def first_dow(year, month, dow):
    day = ((8 + dow) - datetime.date(year, month, 1).weekday()) % 7
    return datetime.date(year, month, day)

class GUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.geometry("500x500")
        self.root.title("Order Processor")

        self.excel_input_frame = tk.Frame(self.root)
        self.excel_file_label = tk.Label(self.excel_input_frame, text="Excel File:")
        self.excel_file_label.pack(padx=10, pady=10)
        self.excel_file_entry = tk.Entry(self.excel_input_frame)
        self.excel_file_entry.pack(padx=10, pady=10)
        self.excel_file_button = tk.Button(self.excel_input_frame, text="Browse Excel", command=self.browse_excel)
        self.excel_file_button.pack(padx=10, pady=10, expand=True)
        self.excel_input_frame.pack(padx=10, pady=10, fill=tk.X)

        self.date_input_frame = tk.Frame(self.root)
        self.date_label = tk.Label(self.date_input_frame, text="Date:")
        self.date_label.pack(padx=10, pady=10)

        self.date_entry = Calendar(self.date_input_frame, selectmode='day')


        self.selected_date_entry = tk.Entry(self.date_input_frame)
        self.selected_date_entry.pack(padx=10, pady=10)
        self.selected_date_entry.insert(0, datetime.date.today().strftime("%d/%m/%Y"))

        self.toggle_button = tk.Button(self.date_input_frame, text="Toggle Calendar", command=self.toggle_calendar)
        self.toggle_button.pack(padx=10, pady=10)

        self.calendar_visible = True  # Flag to track calendar visibility

        self.date_input_frame.pack(padx=10, pady=10, fill=tk.X)
        self.date_entry.pack(padx=10, pady=10)
        self.date_entry.bind("<<CalendarSelected>>", self.update_selected_date)
        self.root.mainloop()

    def browse_excel(self):
        excel_file = filedialog.askopenfilename(filetypes=[("CSV Files (Comma separated)", "*.csv")])
        self.excel_file_entry.delete(0, tk.END)
        self.excel_file_entry.insert(0, excel_file)

    def toggle_calendar(self):
        if self.calendar_visible:
            self.date_entry.pack_forget()
            self.toggle_button.configure(text="Show Calendar")
        else:
            self.date_entry.pack(padx=10, pady=10)
            self.toggle_button.configure(text="Hide Calendar")
        self.calendar_visible = not self.calendar_visible

    def update_selected_date(self, event):
        selected_date = self.date_entry.get_date()
        formatted_date = selected_date
        formatted_date = datetime.datetime.strptime(formatted_date, "%m/%d/%y")
        formatted_date = formatted_date.strftime("%d/%m/%Y")
        self.selected_date_entry.delete(0, tk.END)
        self.selected_date_entry.insert(0, formatted_date)



if __name__ == "__main__":
    gui = GUI()
