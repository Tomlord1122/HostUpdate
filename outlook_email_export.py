import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import win32com.client
import csv
import datetime
import requests

class OutlookEmailExporter:
    def __init__(self, master):
        self.master = master
        master.title("Outlook郵件導出工具")
        master.geometry("600x400")

        # 創建主框架
        main_frame = ttk.Frame(master, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 文本過濾
        ttk.Label(main_frame, text="文本過濾:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.text_filter = ttk.Entry(main_frame, width=40)
        self.text_filter.grid(row=0, column=1, columnspan=2, sticky=tk.W, pady=5)

        # 日期過濾
        ttk.Label(main_frame, text="開始日期:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.start_date = ttk.Entry(main_frame, width=20)
        self.start_date.grid(row=1, column=1, sticky=tk.W, pady=5)
        self.start_date.insert(0, "YYYY-MM-DD")

        ttk.Label(main_frame, text="結束日期:").grid(row=1, column=2, sticky=tk.W, pady=5)
        self.end_date = ttk.Entry(main_frame, width=20)
        self.end_date.grid(row=1, column=3, sticky=tk.W, pady=5)
        self.end_date.insert(0, "YYYY-MM-DD")

        # LLM API URL
        ttk.Label(main_frame, text="LLM API URL:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.llm_api_url = ttk.Entry(main_frame, width=40)
        self.llm_api_url.grid(row=2, column=1, columnspan=2, sticky=tk.W, pady=5)

        # 導出按鈕
        self.export_button = ttk.Button(main_frame, text="導出到CSV", command=self.export_to_csv)
        self.export_button.grid(row=3, column=0, columnspan=2, pady=10)

        # 摘要按鈕
        self.summarize_button = ttk.Button(main_frame, text="摘要選中郵件", command=self.summarize_emails)
        self.summarize_button.grid(row=3, column=2, columnspan=2, pady=10)

        # 郵件列表
        self.email_listbox = tk.Listbox(main_frame, width=80, height=15)
        self.email_listbox.grid(row=4, column=0, columnspan=4, pady=5)

        # 滾動條
        scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=self.email_listbox.yview)
        scrollbar.grid(row=4, column=4, sticky=tk.NS)
        self.email_listbox.config(yscrollcommand=scrollbar.set)

    def export_to_csv(self):
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)
        messages = inbox.Items

        # 應用過濾器
        text_filter = self.text_filter.get()
        start_date = self.start_date.get()
        end_date = self.end_date.get()

        filtered_messages = []
        for message in messages:
            if text_filter and text_filter.lower() not in message.Subject.lower():
                continue
            if start_date != "YYYY-MM-DD":
                message_date = message.ReceivedTime.date()
                if message_date < datetime.datetime.strptime(start_date, "%Y-%m-%d").date():
                    continue
            if end_date != "YYYY-MM-DD":
                message_date = message.ReceivedTime.date()
                if message_date > datetime.datetime.strptime(end_date, "%Y-%m-%d").date():
                    continue
            filtered_messages.append(message)

        # 更新列表框
        self.email_listbox.delete(0, tk.END)
        for message in filtered_messages:
            self.email_listbox.insert(tk.END, f"{message.ReceivedTime.strftime('%Y-%m-%d %H:%M')} - {message.Subject}")

        # 導出到CSV
        file_path = filedialog.asksaveasfilename(defaultextension=".csv")
        if file_path:
            with open(file_path, 'w', newline='', encoding='utf-8') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(["日期", "發件人", "主題", "內容"])
                for message in filtered_messages:
                    writer.writerow([
                        message.ReceivedTime.strftime('%Y-%m-%d %H:%M'),
                        message.SenderName,
                        message.Subject,
                        message.Body
                    ])
            messagebox.showinfo("成功", f"郵件已導出到 {file_path}")

    def summarize_emails(self):
        selected_indices = self.email_listbox.curselection()
        if not selected_indices:
            messagebox.showwarning("警告", "請先選擇要摘要的郵件")
            return

        api_url = self.llm_api_url.get()
        if not api_url:
            messagebox.showwarning("警告", "請輸入有效的LLM API URL")
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
            summary_window.title("郵件摘要")
            summary_window.geometry("500x300")

            summary_text = tk.Text(summary_window, wrap=tk.WORD)
            summary_text.pack(fill=tk.BOTH, expand=True)

            for i, summary in enumerate(summaries):
                summary_text.insert(tk.END, f"郵件 {i+1} 摘要:\n{summary}\n\n")

        except requests.RequestException as e:
            messagebox.showerror("錯誤", f"無法連接到LLM API: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = OutlookEmailExporter(root)
    root.mainloop()