"# mailmerge_app" 

Email Mail Merge Composer Application

Introduction:
-------------------------
The Email Composer Application is a Python-based GUI tool designed to simplify the process of sending personalized emails to multiple recipients using Microsoft Outlook. It allows you to compose emails with either an HTML document or plain text as the email body and provides the flexibility to upload recipient email addresses from various sources like CSV files, Excel files, or Google Sheets.

Functionality:
-------------------------
- Compose personalized emails with HTML or plain text bodies.
- Upload email addresses and additional recipient information from CSV, Excel, or Google Sheets.
- Ensure that the required email fields (Subject, Address to, Body) are filled before sending emails.
- Validate email addresses for proper formatting.
- Display a success message when emails are sent successfully.
- Show the progress of sending emails with a counter indicating the current row being processed.

How to Use:
-------------------------
1. Ensure you have Python installed on your system.

2. Clone or download the "mailmerge_app" repository from GitHub.

https://github.com/Tolonious/mailmerge_app

3. Install the required packages: pip install pandas gspread oauth2client openpyxl

4. Run the "email_app.py" script to launch the Email Composer application:

5. The application's graphical user interface (GUI) will appear.

6. Fill in the necessary information in the respective fields:
- HTML Body File: Browse and select an HTML file to use as the email body (optional if Plain Text is selected).
- Contacts File: Browse and select a CSV, Excel, or Google Sheets file containing recipient information.
- Email Subject: Enter the subject of the email.
- Email Body Type: Select either "HTML" or "Plain Text" to specify the type of email body.
- Email Body: Compose the email body using either HTML tags or plain text (depending on the selected type).

7. Click the "Send Email" button to send the emails to the specified recipients. The application will validate the inputs and display any errors if found.

8. If "HTML" is selected as the Email Body Type, ensure that the HTML Body File is uploaded and exists.

9. If using Google Sheets as the contacts file, a JSON credentials file is required for authentication.

10. When sending emails, a counter will appear at the bottom right of the GUI, indicating the progress (current row number being sent) out of the total rows.

11. Once the emails are sent, a success message will be shown for each recipient.

12. The application will be ready for another email composition once the process is completed.

Contributing:
-------------------------
If you encounter any issues or have suggestions for improvements, feel free to open an issue or submit a pull request on GitHub.

License:
-------------------------
This project is licensed under the MIT License - see the LICENSE file for details.

Author:
-------------------------
Anthony Jackson
