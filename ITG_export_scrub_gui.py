"""
ITG_export_scrubber.py

Author: Josh Smith

Purpose: Stage ITG export data into a single workbook per client with
separate worksheets based on input data csv's. Format cells based on
width of widest string, freeze header row/format as a table,
and zip for sending. Also edit cell data to remove html, unicode issues,
 and leave only readable data.
 """

# imports
import os
import shutil
import csv
import unicodedata
from bs4 import BeautifulSoup
from zipfile import ZipFile
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import scrolledtext
import sys
import logging

# Logging config
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %'
                                                '(levelname)s - %'
                                                '(message)s'
                    )

# Logging trigger
logging.disable(logging.CRITICAL)


class LabelInput(tk.Frame):
    """A widget containing a label and input together.
    Credit: Alan D. Moore "Python GUI Programming with Tkinter"""

    def __init__(
            self, parent, label, var, input_class=ttk.Entry,
            input_args=None, label_args=None, **kwargs
    ):
        super().__init__(parent, **kwargs)
        input_args = input_args or {}
        label_args = label_args or {}
        # The above statements say if label_args or input_args are not None,
        # they are what was passed during init.
        # However, if they are None, then make them empty dicts
        self.variable = var
        self.variable.label_widget = self

        if input_class in (ttk.Checkbutton, ttk.Button):
            input_args["text"] = label
        else:
            self.label = ttk.Label(self, text=label, **label_args)
            self.label.grid(row=0, column=0, sticky=(tk.W + tk.E))

        if input_class in (
            ttk.Checkbutton, ttk.Button, ttk.Radiobutton
        ):
            input_args["variable"] = self.variable
        else:
            input_args["textvariable"] = self.variable

        if input_class == ttk.Radiobutton:
            self.input = tk.Frame(self)
            for v in input_args.pop('values', []):
                button = ttk.Radiobutton(
                    self.input, value=v, text=v, **input_args
                )
                button.pack(
                    side=tk.LEFT, ipadx=10, ipady=2, expand=True, fill='x'
                )
        else:
            self.input = input_class(self, **input_args)

        self.input.grid(row=1, column=0, sticky=(tk.E + tk.W))
        self.columnconfigure(0, weight=1)

    def grid(self, sticky=(tk.E + tk.W), **kwargs):
        """Override grid to add default sticky values"""
        super().grid(sticky=sticky, **kwargs)


class AppPage(ttk.Frame):
    """Application page class from which all other pages will inherit."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._vars = {}

    def _add_frame(self, label, cols=2):
        """Add a label frame to the form
        Credit: Alan D. Moore "Python GUI Programming with Tkinter"""

        frame = ttk.LabelFrame(self, text=label)
        frame.grid(sticky=tk.W + tk.E)
        for i in range(cols):
            frame.columnconfigure(i, weight=1)
        return frame

    def get(self):
        data = dict()
        for key, variable in self._vars.items():
            try:
                data[key] = variable.get()
            except tk.TclError:
                message = f'Error in field: {key}. Data was not saved!'
                raise ValueError(message)
        return data


class MainPage(AppPage):
    """Main Page to select options, change folder, and run"""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._vars = {'Batch Size': tk.StringVar(),
                      'Post Job': tk.StringVar(),
                      'Zip?': tk.StringVar(),
                    }
        size_default = self._add_frame('Processing single file or a folder?')
        post_default = self._add_frame(
            'Delete or Keep original export when finished?'
        )
        zip_default = self._add_frame('Zip the output when finished?')
        buttons = self._add_frame('')

        LabelInput(size_default, '', input_class=ttk.Radiobutton,
                   var=self._vars['Batch Size'],
                   input_args={'values': ['Single File', 'Folder']}
                   ).grid(row=0, column=0, sticky=(tk.W + tk.E))

        LabelInput(post_default, '', input_class=ttk.Radiobutton,
                   var=self._vars['Post Job'],
                   input_args={'values': ['Delete', 'Keep']}
                   ).grid(row=1, column=0, sticky=(tk.W + tk.E))

        LabelInput(zip_default, '', input_class=ttk.Radiobutton,
                   var=self._vars['Zip?'],
                   input_args={'values': ['Yes', 'No']}
                   ).grid(row=2, column=0, sticky=(tk.W + tk.E))

        self.run_button = tk.Button(buttons, text='Run',
                                    command=self._on_run
                                    )
        self.run_button.grid(row=2, column=0, sticky='ew')

        self.select_target = tk.Button(buttons, text='Select Target',
                                       command=self._on_target
                                       )
        self.select_target.grid(row=1, column=0, sticky='ew')

        self.quit_button = tk.Button(
            buttons,
            text='Quit',
            command=self._on_quit
        )
        self.quit_button.grid(row=3, column=0, sticky='ew')

    @staticmethod
    def _on_quit():
        """Command to exit program"""
        sys.exit()

    @staticmethod
    def _on_run():
        """Command to exit program"""
        pass

    @staticmethod
    def _on_target():
        """Command to exit program"""
        pass


class Application(tk.Tk):
    """Application root window"""
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.status = tk.StringVar()
        ttk.Label(
            self, textvariable=self.status
        ).grid(sticky=(tk.W + tk.E), row=2, padx=10)
        self.m_page = ''
        self.main_label = ''
        self.title("ITG Export Scrubber 1.0")
        self.minsize(400, 300)
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
    app = Application()
    app.grid_columnconfigure(0, weight=1)
    app.mainloop()