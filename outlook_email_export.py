import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import win32com.client
import csv
import datetime
import requests
from tkcalendar import DateEntry

class OutlookEmailExporter:
    def __init__(self, master):
        self.master = master
        master.title("Outlook Email Export Tool")
        master.geometry("600x400")

        # Create main frame
        main_frame = ttk.Frame(master, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Text filter
        ttk.Label(main_frame, text="Text Filter:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.text_filter = ttk.Entry(main_frame, width=40)
        self.text_filter.grid(row=0, column=1, columnspan=2, sticky=tk.W, pady=5)

        # Date filter
        ttk.Label(main_frame, text="Start Date:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.start_date = DateEntry(main_frame, width=20, date_pattern='yyyy-mm-dd')
        self.start_date.grid(row=1, column=1, sticky=tk.W, pady=5)

        ttk.Label(main_frame, text="End Date:").grid(row=1, column=2, sticky=tk.W, pady=5)
        self.end_date = DateEntry(main_frame, width=20, date_pattern='yyyy-mm-dd')
        self.end_date.grid(row=1, column=3, sticky=tk.W, pady=5)

        # LLM API URL
        ttk.Label(main_frame, text="LLM API URL:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.llm_api_url = ttk.Entry(main_frame, width=40)
        self.llm_api_url.grid(row=2, column=1, columnspan=2, sticky=tk.W, pady=5)

        # Export button
        self.export_button = ttk.Button(main_frame, text="Export to CSV", command=self.export_to_csv)
        self.export_button.grid(row=3, column=0, columnspan=2, pady=10)

        # Summarize button
        self.summarize_button = ttk.Button(main_frame, text="Summarize Selected Emails", command=self.summarize_emails)
        self.summarize_button.grid(row=3, column=2, columnspan=2, pady=10)

        # Email list
        self.email_listbox = tk.Listbox(main_frame, width=80, height=15)
        self.email_listbox.grid(row=4, column=0, columnspan=4, pady=5)

        # Scrollbar
        scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=self.email_listbox.yview)
        scrollbar.grid(row=4, column=4, sticky=tk.NS)
        self.email_listbox.config(yscrollcommand=scrollbar.set)

    def export_to_csv(self):
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)
        messages = inbox.Items

        # Apply filters
        text_filter = self.text_filter.get()
        start_date = self.start_date.get_date()
        end_date = self.end_date.get_date()

        filtered_messages = []
        for message in messages:
            if text_filter and text_filter.lower() not in message.Subject.lower():
                continue
            message_date = message.ReceivedTime.date()
            if start_date and message_date < start_date:
                continue
            if end_date and message_date > end_date:
                continue
            filtered_messages.append(message)

        # Update listbox
        self.email_listbox.delete(0, tk.END)
        for message in filtered_messages:
            self.email_listbox.insert(tk.END, f"{message.ReceivedTime.strftime('%Y-%m-%d %H:%M')} - {message.Subject}")

        # Export to CSV
        file_path = filedialog.asksaveasfilename(defaultextension=".csv")
        if file_path:
            with open(file_path, 'w', newline='', encoding='utf-8') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(["Date", "Sender", "Subject", "Content"])
                for message in filtered_messages:
                    writer.writerow([
                        message.ReceivedTime.strftime('%Y-%m-%d %H:%M'),
                        message.SenderName,
                        message.Subject,
                        message.Body
                    ])
            messagebox.showinfo("Success", f"Emails exported to {file_path}")

    def summarize_emails(self):
        selected_indices = self.email_listbox.curselection()
        if not selected_indices:
            messagebox.showwarning("Warning", "Please select emails to summarize")
            return

        api_url = self.llm_api_url.get()
        if not api_url:
            messagebox.showwarning("Warning", "Please enter a valid LLM API URL")
            return

        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)
        messages = inbox.Items

        selected_messages = [messages[int(index)] for index in selected_indices]
        email_contents = [f"Subject: {msg.Subject}\n\nBody: {msg.Body}" for msg in selected_messages]

        try:
            response = requests.post(api_url, json={"texts": email_contents})
            response.raise_for_status()
            summaries = response.json()["summaries"]

            summary_window = tk.Toplevel(self.master)
            summary_window.title("Email Summaries")
            summary_window.geometry("500x300")

            summary_text = tk.Text(summary_window, wrap=tk.WORD)
            summary_text.pack(fill=tk.BOTH, expand=True)

            for i, summary in enumerate(summaries):
                summary_text.insert(tk.END, f"Email {i+1} Summary:\n{summary}\n\n")

        except requests.RequestException as e:
            messagebox.showerror("Error", f"Unable to connect to LLM API: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = OutlookEmailExporter(root)
    root.mainloop()