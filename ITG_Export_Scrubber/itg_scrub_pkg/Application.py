import tkinter as tk
from tkinter import ttk
from ITG_Export_Scrubber.itg_scrub_pkg.MainPage import MainPage


class Application(tk.Tk):
    """Application root window"""
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.m_page = ''
        self.main_label = ''
        self.title(" TPG ITG Export Scrubber v1.66")
        self.minsize(400, 350)
        self.main_page()

    def main_page(self):
        self.m_page = MainPage(self)
        self.main_label = ttk.Label(
            self,
            text="TPG ITG Export Scrubber",
            font=("TKDefaultFont", 14))
        self.main_label.grid(row=0)
        self.m_page.grid(row=1, padx=10, sticky=(tk.W + tk.E))


if __name__ == "__main__":
    pass