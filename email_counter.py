import tkinter as tk

class EmailCounter(tk.Label):
    def __init__(self, root, total_rows):
        super().__init__(root, text="", anchor="se")
        self.total_rows = total_rows
        self.current_row = 0
        self.update_text()
    
    def update_row(self, current_row):
        self.current_row = current_row
        self.update_text()
    
    def update_text(self):
        text = f"Sending email {self.current_row} of {self.total_rows}"
        self.config(text=text)

