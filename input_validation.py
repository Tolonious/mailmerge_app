import re

def is_valid_email(email):
    # Split email addresses by ";" if present
    email_list = email.split(";")

    for email_address in email_list:
        # Check the validity of each email address
        if not re.match(r"[^@]+@[^@]+\.[^@]+", email_address):
            return False

    return True
