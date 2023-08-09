import webbrowser
from bs4 import BeautifulSoup
import customtkinter as ctk

def preview_email(subject, body, to_addresses, cc_addresses, bcc_addresses):
    # Compose the email content
    email_content = f"Subject: {subject}\nTo: {', '.join(to_addresses)}"
    if cc_addresses:
        email_content += f"\nCC: {', '.join(cc_addresses)}"
    if bcc_addresses:
        email_content += f"\nBCC: {', '.join(bcc_addresses)}"
    email_content += f"\n\n{body}"

    # Open a new window to display the email preview
    preview_window = ctk.CTkToplevel()
    preview_window.title("Email Preview")
    preview_window.geometry("600x400")

    # Display the email content in a Text widget
    text_widget = ctk.CTkTextbox(preview_window, wrap=ctk.WORD, width=80, height=20)
    text_widget.insert(ctk.END, email_content)
    text_widget.config(state=ctk.DISABLED)  # Make the Text widget read-only
    text_widget.pack(padx=10, pady=10)

    # Add a 'Send Email' button to close the preview window and proceed with sending the email
    ctk.CTkButton(preview_window, text="Send Email", command=preview_window.destroy).pack(pady=5)

    preview_window.mainloop()
