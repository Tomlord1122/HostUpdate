import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
import datetime
import win32com.client

class App:
  
    
    def __init__(self, master):
        self.master = master
        self.master.title("Outlook Collector")
        self.master.geometry("1024x640")  # width x height
        self.master.config(bg="#323232")
        
        # Connect to Outlook
        self.outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        self.inbox = self.outlook.GetDefaultFolder(6)
        

        # Collect Button
        self.btn = tk.Button(self.master, text="Collect", command=self.collect_emails)
        self.btn.config(bg="#323232", fg="black", font=("Arial", 14), width=10, height=2)
        self.btn.grid(row=0, column=0, padx=10, pady=10)
        
        # Export Button
        self.export_btn = tk.Button(self.master, text="Export", command=self.export_emails)
        self.export_btn.config(bg="#323232", fg="black", font=("Arial", 14), width=10, height=2)
        self.export_btn.grid(row=1, column=0, padx=10, pady=10)
        
        # Date Entry
        self.from_date_entry = DateEntry(self.master, width=12,foreground='black', borderwidth=2)
        self.from_date_entry.grid(row=0, column=2, padx=10, pady=10)
        self.from_date_label = tk.Label(self.master, text="From Date", font=("Arial", 14), bg="#323232", fg="white")
        self.from_date_label.grid(row=0, column=1, padx=5, pady=20)
        self.to_date_entry = DateEntry(self.master, width=12, foreground='black', borderwidth=2)
        self.to_date_entry.grid(row=1, column=2, padx=10, pady=10)
        self.to_date_label = tk.Label(self.master, text="To Date", font=("Arial", 14), bg="#323232", fg="white")
        self.to_date_label.grid(row=1, column=1, padx=5, pady=20)
        
        
        # Email List
        self.email_list = tk.Listbox(self.master, width=100, height=30)
        self.email_list.place(x=10, y=150)
        

    def collect_emails(self):
        # Placeholder for the collect function
        print("Collecting emails...")
        self.to_date_value = self.to_date_entry.get_date()
        self.from_date_value = self.from_date_entry.get_date()
        print(f"From Date: {self.from_date_value}")
        print(f"To Date: {self.to_date_value}")
    
    def export_emails(self):
        # Placeholder for the export function
        print("Exporting emails...")
        
    def load_emails(self):
        today = datetime.date.today() # fethe the current date
        messages = self.inbox.Items # get all the items in the inbox
        messages.Sort("[ReceivedTime]", True) # sort the items by the received time in ascending order
        messages = messages.Restrict("[ReceivedTime] >= '" + today.strftime('%m/%d/%Y') + "'") # get all the items that were received after the current date
        
        for message in messages:
            self.email_list.insert(tk.END, message.Subject)
            
            
            
    def filter_emails(self):
        pass
        

    
# Application
if __name__ == "__main__":
    win = tk.Tk()
    app = App(win)
    win.mainloop()