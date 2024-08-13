import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
import win32com.client
from openai import OpenAI
import os
from dotenv import load_dotenv


load_dotenv()

client = OpenAI(
  api_key=os.getenv("OPENAI_API_KEY")
)

input_map = {
    'Select Filter': "",
    'adk': 'adk',
    'polycam': 'polycam'
}

class App:
    def __init__(self, master):
        self.master = master
        self.master.title("Outlook Collector")
        self.master.geometry("770x640")  # width x height
        self.master.config(bg="#323232")
        self.master.iconbitmap("hp.ico")
    
        # Connect to Outlook
        self.outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        self.inbox = self.outlook.GetDefaultFolder(6)

        
        
        # Collect Button
        self.btn = tk.Button(self.master, text="Collect", command=self.collect_emails)
        self.btn.config(bg="white", fg="black", font=("Microsoft Sans Serif", 10), width=10, height=2)
        self.btn.grid(row=0, column=0, padx=10, pady=10)
        
        # Export Button
        self.export_btn = tk.Button(self.master, text="Summarize", command=self.summarize_mail)
        self.export_btn.config(bg="white", fg="black", font=("Microsoft Sans Serif", 10), width=10, height=2)
        self.export_btn.grid(row=1, column=0, padx=10, pady=10)

        # Filter input label
        self.filter_label = tk.Label(self.master, text="Filter:", font=("Microsoft Sans Serif", 10), bg="#323232", fg="white")
        self.filter_label.grid(row=0, column=3, padx=10, pady=10)
        
        self.filter_input = ttk.Combobox(self.master, width=30)
        self.filter_input['values'] = ['Select Filter', 'adk', 'polycam']
        self.filter_input.current(0)
        self.filter_input.grid(row=0, column=4, padx=10, pady=10)
        
        # Date Entry
        self.from_date_entry = DateEntry(self.master, width=12,foreground='black', borderwidth=2)
        self.from_date_entry.grid(row=0, column=2, padx=10, pady=10)
        self.from_date_label = tk.Label(self.master, text="From Date", font=("Microsoft Sans Serif", 10), bg="#323232", fg="white")
        self.from_date_label.grid(row=0, column=1, padx=5, pady=20)
        self.to_date_entry = DateEntry(self.master, width=12, foreground='black', borderwidth=2)
        self.to_date_entry.grid(row=1, column=2, padx=10, pady=10)
        self.to_date_label = tk.Label(self.master, text="To Date", font=("Microsoft Sans Serif", 10), bg="#323232", fg="white")
        self.to_date_label.grid(row=1, column=1, padx=5, pady=20)
        


        
        # Email List
        frame = ttk.Frame(self.master)
        frame.place(x=10, y=150)

        self.email_list = tk.Listbox(frame, width=50, height=30)
        self.email_list.pack(side=tk.LEFT, padx=10)

        self.summary_box = tk.Text(frame, width=50, height=30, font=("Microsoft Sans Serif", 10))
        self.summary_box.pack(side=tk.LEFT, padx=10)

        self.collect_emails()

         

    def collect_emails(self):
        # Placeholder for the collect function
        self.email_list.delete(0, tk.END)
        self.mail_body = []
        print("Collecting emails...")
        to_date = self.to_date_entry.get_date()
        from_date = self.from_date_entry.get_date()
        print(from_date, to_date)
        messages = self.inbox.Items # get all the items in the inbox
        messages.Sort("[ReceivedTime]", True) # sort the items by the received time in ascending order
        messages = messages.Restrict("[ReceivedTime] >= '" + from_date.strftime('%m/%d/%Y') + "' AND [ReceivedTime] <= '" + to_date.strftime('%m/%d/%Y') + "'") # get all the items that were received between the specified dates
        filter_input = self.filter_input.get()
        filter = input_map[filter_input]
        for message in messages:
            # message_cmp = message.Sender.GetExchangeUser().PrimarySmtpAddress.lower()
            message_cmp = message.Subject.lower()
            print(message_cmp)
            if filter in message_cmp:
                self.email_list.insert(tk.END, message.Subject)
                self.mail_body.append(message.Body)
        print(len(self.mail_body))
            
    def summarize_mail(self):
        print("Summarizing emails...")
        self.summary_box.delete(1.0, tk.END) # clear the summary box
        self.email_list.delete(0, tk.END)
        self.email_list.insert(tk.END, "Summarizing emails...")
        # 將郵件內容合併為一個字符串
        all_emails = "\n\n".join(self.mail_body)
        
        try:
            stream = client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {"role": "system", "content": "You are a system engineer familiar with Windows and the Assessment and Deployment Kit (ADK)."},
                    {"role": "user", "content": f"Categorize and summarize the problems encountered by ODMs in the following emails, make sure you classify the ODMs:\n\n{all_emails}"}
                ],
                stream=True,
            )
            
            for chunk in stream:
                if chunk.choices[0].delta.content is not None:
                    content = chunk.choices[0].delta.content
                    # self.email_list.insert(tk.END, content)
                    print(content, end="")
                    self.summary_box.insert(tk.END, content)
            print("\nSummarization completed.")
            self.mail_body = []
        except Exception as e:
            print(f"An error occurred: {e}")
            self.email_list.insert(tk.END, f"Error: {e}")
        

        

    
# Application
if __name__ == "__main__":
    win = tk.Tk()
    app = App(win)
    win.mainloop()