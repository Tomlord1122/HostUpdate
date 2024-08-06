import tkinter as tk
from tkinter import ttk, filedialog
import win32com.client
import csv
import datetime
import os
import openai
from tkcalendar import DateEntry

class OutlookExportApp:
    def __init__(self, master):
        self.master = master
        master.title("Outlook Email Export Application")
        
        # Set up UI elements
        self.setup_ui()
        
        # Connect to Outlook
        self.outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        self.inbox = self.outlook.GetDefaultFolder(6)
        
        # Get today's emails
        self.load_today_emails()

    def setup_ui(self):
        # Create main frame
        main_frame = ttk.Frame(self.master, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Text filter
        ttk.Label(main_frame, text="Text Filter:").grid(row=0, column=0, sticky=tk.W)
        self.text_filter = ttk.Entry(main_frame, width=30)
        self.text_filter.grid(row=0, column=1, sticky=(tk.W, tk.E))
        
        # 日期過濾器
        ttk.Label(main_frame, text="日期過濾器:").grid(row=1, column=0, sticky=tk.W)
        self.date_filter = DateEntry(main_frame, width=30, date_pattern='yyyy-mm-dd')
        self.date_filter.grid(row=1, column=1, sticky=(tk.W, tk.E))
        
        # Filter button
        ttk.Button(main_frame, text="Filter", command=self.filter_emails).grid(row=2, column=0, columnspan=2)
        
        # Email list
        self.email_list = tk.Listbox(main_frame, width=50, height=10)
        self.email_list.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E))
        
        # Export button
        ttk.Button(main_frame, text="Export", command=self.export_emails).grid(row=4, column=0)
        
        # Summarize button
        ttk.Button(main_frame, text="Summarize", command=self.summarize_email).grid(row=4, column=1)

    def load_today_emails(self):
        today = datetime.date.today()
        messages = self.inbox.Items
        messages.Sort("[ReceivedTime]", True)
        messages = messages.Restrict("[ReceivedTime] >= '" + today.strftime('%m/%d/%Y') + "'")
        
        for message in messages:
            self.email_list.insert(tk.END, message.Subject)

    def filter_emails(self):
        text = self.text_filter.get().lower()
        date_str = self.date_filter.get()
        
        self.email_list.delete(0, tk.END)
        
        messages = self.inbox.Items
        messages.Sort("[ReceivedTime]", True)
        
        if date_str:
            try:
                filter_date = datetime.datetime.strptime(date_str, '%Y-%m-%d').date()
                messages = messages.Restrict("[ReceivedTime] >= '" + filter_date.strftime('%m/%d/%Y') + "'")
            except ValueError:
                tk.messagebox.showerror("Error", "Invalid date format. Please use YYYY-MM-DD format.")
                return
        
        for message in messages:
            if text.lower() in message.Subject.lower() or text.lower() in message.Body.lower():
                self.email_list.insert(tk.END, message.Subject)

    def export_emails(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".csv")
        if not file_path:
            return
        
        with open(file_path, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(['Subject', 'Sender', 'Received Time', 'Body'])
            
            for index in self.email_list.curselection():
                subject = self.email_list.get(index)
                message = self.inbox.Items.Find("[Subject] = '" + subject + "'")
                writer.writerow([message.Subject, message.SenderName, message.ReceivedTime, message.Body])
        
        tk.messagebox.showinfo("Success", f"Emails exported to {file_path}")

    def summarize_email(self):
        selected_indices = self.email_list.curselection()
        if not selected_indices:
            tk.messagebox.showwarning("Warning", "Please select an email first")
            return
        
        subject = self.email_list.get(selected_indices[0])
        message = self.inbox.Items.Find("[Subject] = '" + subject + "'")
        
        openai.api_key = os.getenv("OPENAI_API_KEY")
        if not openai.api_key:
            tk.messagebox.showerror("Error", "OPENAI_API_KEY environment variable not set")
            return
        
        try:
            response = openai.Completion.create(
                engine="text-davinci-002",
                prompt=f"Please summarize the following email content:\n\n{message.Body}",
                max_tokens=150
            )
            summary = response.choices[0].text.strip()
            tk.messagebox.showinfo("Email Summary", summary)
        except Exception as e:
            tk.messagebox.showerror("Error", f"Unable to generate summary: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = OutlookExportApp(root)
    root.mainloop()