import os
import pandas as pd
import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from email_utils import send_email
from input_validation import is_valid_email
from email_counter import EmailCounter
from email_preview import preview_email

# Import required libraries for HTML preview
import webbrowser
from bs4 import BeautifulSoup

class EmailApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Email Composer")
        self.root.geometry('640x500')
        self.root.minsize(width=620, height=460)
        ctk.set_appearance_mode('system')
        ctk.set_default_color_theme('green')

        # Variables to store user inputs
        self.contacts_file_path = ctk.StringVar()
        self.email_body_type = ctk.StringVar(value="Choose Format")
        self.html_file_path = ctk.StringVar()
        self.email_subject = ctk.StringVar()
        self.email_body = ctk.StringVar()
        self.body_format = ctk.StringVar()
        
        # Create and Configureure widgets
        self.create_widgets()

    def create_widgets(self):
        # configure font
        my_font = ctk.CTkFont( family='Arial', size= 11, weight= 'normal')
        
        # Row 0: Contact File Input
        ctk.CTkLabel(self.root, font= my_font.configure(weight= 'bold'), text="Contact file:").grid(row=0, column=0, padx=20, pady=5, sticky='w')
        ctk.CTkLabel(self.root, font= my_font, textvariable=self.contacts_file_path).grid(row=0, column=1, columnspan=2, padx=5, pady=5, sticky='w')
        ctk.CTkButton(self.root, font= my_font, text="Browse", command=self.browse_contacts_file).grid(row=0, column=4, padx=5, pady=5)

        # Row 1: Email Body Type
        ctk.CTkLabel(self.root, font= my_font, text="Email Body Type:").grid(row=1, column=0, padx=20, pady=5)
        ctk.CTkComboBox(self.root, font= my_font, values=["Plain Text", 'HMTL'], command=self.body_format, variable=self.email_body_type).grid(row=1, column=1, padx=5, pady=5)

        # Row 2: HTML File Input (Visible when 'HTML' is chosen)
        body_format = self.body_format.get()
        if body_format == "Plain Text":
            self.browse_html_file = 'disabled'
        ctk.CTkLabel(self.root, font= my_font, text="HTML File:").grid(row=2, column=0, padx=20, pady=5, sticky= 'w')
        ctk.CTkLabel(self.root, font= my_font, textvariable=self.html_file_path).grid(row=2, column=1, columnspan=2, padx=5, pady=5, sticky='w')
        ctk.CTkButton(self.root, font= my_font, text="Browse", command=self.browse_html_file).grid(row=2, column=4, padx=5, pady=5)

        # Rows 4 + 5: Email Subject and Body
        self.frame_subject_body = ctk.CTkFrame(self.root)
        self.frame_subject_body.grid(row=4, column=0, columnspan=5, padx=15, pady=15, sticky="nsew")
        ctk.CTkLabel(self.frame_subject_body, font= my_font, text="Subject:").grid(row=0, column=0, padx=5, pady=5, sticky= 'w')
        ctk.CTkEntry(self.frame_subject_body, font= my_font, textvariable=self.email_subject, width=55).grid(row=0, column=2, columnspan=2, padx=5, pady=5, sticky= 'ew')

        ctk.CTkLabel(self.frame_subject_body, font= my_font, text="Body:").grid(row=1, column=0, padx=5, pady=5, sticky= 'w')
        self.body_text = ctk.CTkTextbox(self.frame_subject_body, wrap=ctk.WORD, width=60, height=10)
        self.body_text.grid(row=1, column=2, rowspan=5,columnspan=4,padx=5, pady=5, sticky="nsew")

        # Row 6: Preview and Send Buttons
        ctk.CTkButton(self.root, font= my_font, text="Preview", command=self.preview_email).grid(row=6, column=3, padx=5, pady=10)
        ctk.CTkButton(self.root, font= my_font, text="Send", command=self.send_email).grid(row=6, column=4, padx=5, pady=10)

        # Row 7: Error Field (Hidden by default)
        self.error_label = ctk.CTkLabel(self.root, font= my_font, text="", fg_color="red")
        self.error_label.grid(row=7, column=0, columnspan=3, padx=5, pady=5, sticky="se")
        self.hide_error_label()
        
    def browse_html_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("HTML Files", "*.html")])
        if file_path:
            file_name = os.path.basename(file_path)
            self.html_file_path.set(file_name)

    def browse_contacts_file(self):
        file_path = filedialog.askopenfilename(filetypes=[
            ("CSV Files", "*.csv"),
            ("Excel Files", "*.xls;*.xlsx"),
            ("Google Sheets", "*.gsheet")
        ])
        if file_path:
            file_name = os.path.basename(file_path)
            self.contacts_file_path.set(file_name)

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
        self.body_text.delete(1.0, ctk.END)
        self.body_text.insert(ctk.END, soup.prettify())
    
    def preview_email(self):
        # Get the email details from the input fields
        subject = self.email_subject.get()
        email_body_type = self.email_body_type.get()
        email_body = self.body_text.get(1.0, tk.END).strip()

        if not subject:
            self.display_error("Email subject is empty.")
            return

        if not email_body:
            self.display_error("Email body is empty.")
            return

        # Get the email addresses from the input fields (handling multiple addresses separated by ";")
        to_addresses = self.to_address_entry.get().split(";")
        cc_addresses = self.cc_address_entry.get().split(";") if self.cc_address_entry.get() else []
        bcc_addresses = self.bcc_address_entry.get().split(";") if self.bcc_address_entry.get() else []

        # Validate the email addresses
        invalid_addresses = []
        for address_list, address_type in [(to_addresses, "'To'"), (cc_addresses, "'CC'"), (bcc_addresses, "'BCC'")]:
            for address in address_list:
                if address and not is_valid_email(address):
                    invalid_addresses.append(f"Invalid {address_type} email format: {address}")

        if invalid_addresses:
            error_message = "\n".join(invalid_addresses)
            self.display_error(error_message)
            return

        # If the email body type is 'HTML', read the HTML file content
        if email_body_type == "HTML":
            html_file_path = self.html_file_path.get()
            if not os.path.isfile(html_file_path):
                self.display_error("The HTML file does not exist.")
                return

            with open(html_file_path, "r", encoding="utf-8") as html_file:
                email_body = html_file.read()

        # Show the email preview window
        preview_email(subject, email_body, to_addresses, cc_addresses, bcc_addresses)

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
                self.hide_error_label()
            except Exception as e:
                self.display_error(f"Error sending email to '{to_address}': {str(e)}")

    def display_error(self, message):
        self.error_label.set_text(message)
        self.error_label.grid()  # Show the error message label
        self.root.update()  # Update the GUI to make the label visible

    def hide_error_label(self):
        self.error_label.grid_remove() 

if __name__ == "__main__":
    root = ctk.CTk()
    app = EmailApp(root)
    root.mainloop()