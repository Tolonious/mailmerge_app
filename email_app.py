import os
import webbrowser
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from email_utils import send_email
from input_validation import is_valid_email
from email_counter import EmailCounter  
from bs4 import BeautifulSoup

class EmailApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Email Composer")
        
        # Variables to store user inputs
        self.html_file_path = tk.StringVar()
        self.contacts_file_path = tk.StringVar()
        self.email_subject = tk.StringVar()
        self.email_body_type = tk.StringVar(value="HTML")

        # Create and configure widgets
        self.create_widgets()

    def create_widgets(self):
        # HTML body file input
        tk.Label(self.root, text="HTML Body File:").grid(row=0, column=0, padx=5, pady=5)
        self.html_entry = tk.Entry(self.root, textvariable=self.html_file_path, width=50)
        self.html_entry.grid(row=0, column=1, padx=5, pady=5)
        tk.Button(self.root, text="Browse", command=self.browse_html_file).grid(row=0, column=2, padx=5, pady=5)

        # Contacts file input
        tk.Label(self.root, text="Contacts File:").grid(row=1, column=0, padx=5, pady=5)
        self.contacts_entry = tk.Entry(self.root, textvariable=self.contacts_file_path, width=50)
        self.contacts_entry.grid(row=1, column=1, padx=5, pady=5)
        tk.Button(self.root, text="Browse", command=self.browse_contacts_file).grid(row=1, column=2, padx=5, pady=5)

        # Email subject input
        tk.Label(self.root, text="Email Subject:").grid(row=2, column=0, padx=5, pady=5)
        self.subject_entry = tk.Entry(self.root, textvariable=self.email_subject, width=50)
        self.subject_entry.grid(row=2, column=1, padx=5, pady=5)

        # Email body type
        tk.Label(self.root, text="Email Body Type:").grid(row=3, column=0, padx=5, pady=5)
        tk.Radiobutton(self.root, text="HTML", variable=self.email_body_type, value="HTML").grid(row=3, column=1, padx=5, pady=5)
        tk.Radiobutton(self.root, text="Plain Text", variable=self.email_body_type, value="Text").grid(row=3, column=2, padx=5, pady=5)

        # Email body input
        tk.Label(self.root, text="Email Body:").grid(row=4, column=0, padx=5, pady=5)
        self.body_text = tk.Text(self.root, wrap=tk.WORD, width=60, height=10)
        self.body_text.grid(row=4, column=1, columnspan=2, padx=5, pady=5)

       # Preview HTML file button
        tk.Button(self.root, text="Preview HTML", command=self.preview_html_file).grid(row=5, column=0, padx=5, pady=10)

        # Send email button
        tk.Button(self.root, text="Send Email", command=self.send_email).grid(row=5, column=1, columnspan=2, padx=5, pady=10)

        # Error message display
        self.error_text = tk.Text(self.root, wrap=tk.WORD, width=60, height=8)
        self.error_text.grid(row=6, column=0, columnspan=3, padx=5, pady=5)

    def browse_html_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("HTML Files", "*.html")])
        if file_path:
            self.html_file_path.set(file_path)

    def browse_contacts_file(self):
        file_path = filedialog.askopenfilename(filetypes=[
            ("CSV Files", "*.csv"),
            ("Excel Files", "*.xls;*.xlsx"),
            ("Google Sheets", "*.gsheet")
        ])
        if file_path:
            self.contacts_file_path.set(file_path)

    def read_contacts_from_file(self):
        contacts_file = self.contacts_file_path.get()

        if contacts_file.endswith('.csv'):
            df = pd.read_csv(contacts_file)
        elif contacts_file.endswith('.xls') or contacts_file.endswith('.xlsx'):
            df = pd.read_excel(contacts_file)
        elif contacts_file.endswith('.gsheet'):
            credentials_file = filedialog.askopenfilename(filetypes=[("JSON Files", "*.json")])
            if not credentials_file:
                self.display_error("Google Sheets credentials file is required.")
                return

            scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
            credentials = ServiceAccountCredentials.from_json_keyfile_name(credentials_file, scope)
            gc = gspread.authorize(credentials)
            sheet_name = os.path.basename(contacts_file).split('.')[0]
            try:
                worksheet = gc.open(sheet_name).sheet1
                records = worksheet.get_all_records()
                df = pd.DataFrame(records)
            except gspread.exceptions.APIError:
                self.display_error("Unable to access the Google Sheet.")
                return
        else:
            self.display_error("Unsupported contacts file format.")
            return None

        return df

    def preview_html_file(self):
        html_file = self.html_file_path.get()
        if not os.path.isfile(html_file):
            self.display_error("The HTML file does not exist.")
            return

        # Read the HTML file and show the preview in the Email body field
        with open(html_file, "r") as file:
            html_content = file.read()

        soup = BeautifulSoup(html_content, "html.parser")
        self.body_text.delete(1.0, tk.END)
        self.body_text.insert(tk.END, soup.prettify())
    
    def send_email(self):
        html_file = self.html_file_path.get()
        subject = self.email_subject.get()
        email_body_type = self.email_body_type.get()

        if email_body_type == "HTML":
            # If HTML option is chosen, read the HTML file and use it as the email body
            if not os.path.isfile(html_file):
                self.display_error("The HTML file does not exist.")
                return

            with open(html_file, "r") as file:
                email_body = file.read()
        else:
            # If Plain Text option is chosen, use the text in the Email body field
            email_body = self.body_text.get(1.0, tk.END).strip()
        
        # Split the email addresses by ";" and pass them as a list
        to_addresses = self.to_address_entry.get().split(";")
        cc_addresses = self.cc_address_entry.get().split(";")
        bcc_addresses = self.bcc_address_entry.get().split(";")

        df = self.read_contacts_from_file()
        if df is None:
            return
        
        total_rows = len(df)
        email_counter = EmailCounter(self.root, total_rows)  # Create an instance of EmailCounter

        for index, row in df.iterrows():
            to_address = row["Address to"]
            cc_address = row["CC"]
            bcc_address = row["BCC"]

            if not to_address:
                self.display_error("'Address to' field is empty.")
                continue

            if not subject:
                self.display_error("Email subject is empty.")
                continue

            if not email_body:
                self.display_error("Email body is empty.")
                continue

            if not is_valid_email(to_address):
                self.display_error(f"Invalid 'Address to' email format: {to_address}")
                continue

            if cc_address and not is_valid_email(cc_address):
                self.display_error(f"Invalid 'CC' email format: {cc_address}")
                continue

            if bcc_address and not is_valid_email(bcc_address):
                self.display_error(f"Invalid 'BCC' email format: {bcc_address}")
                continue

            try:
                 # Update the EmailCounter with the current row number
                email_counter.update_row(index + 1)

                # Send email for each contact row
                send_email(subject, email_body, to_address, cc=cc_address, bcc=bcc_address)
                messagebox.showinfo("Success", f"Email sent to: {to_address}")
            except Exception as e:
                self.display_error(f"Error sending email to '{to_address}': {str(e)}")

    def display_error(self, message):
        self.error_text.delete(1.0, tk.END)
        self.error_text.insert(tk.END, message)

if __name__ == "__main__":
    root = tk.Tk()
    app = EmailApp(root)
    root.mainloop()
