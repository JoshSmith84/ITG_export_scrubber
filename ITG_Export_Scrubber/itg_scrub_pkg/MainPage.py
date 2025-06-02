from ITG_Export_Scrubber.itg_scrub_pkg.widgets import AppPage, LabelInput
from ITG_Export_Scrubber.itg_scrub_pkg.process_exports import process_exports
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import os
import sys
import datetime


class MainPage(AppPage):
    """Main Page to select options, change folder, and run"""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._vars = {'Batch Size': tk.StringVar(None, 'Folder'),
                      'Post Job': tk.StringVar(None, 'Delete'),
                      'Zip?': tk.StringVar(None, 'No'),
                      }

        self.input_folder = ''
        self.input_file = ''
        self.err_present = 0
        self.err_count = 0

        # Initialize Main Page GUI
        size_default = self._add_frame(
            'Processing a folder or single file?'
        )
        post_default = self._add_frame(
            'Delete or Keep original export when finished?'
        )
        zip_default = self._add_frame('Zip the output when finished?')
        buttons = self._add_frame('')

        LabelInput(size_default, '', input_class=ttk.Radiobutton,
                   var=self._vars['Batch Size'],
                   input_args={'values': ['Folder', 'Single File']}
                   ).grid(row=0, column=0, sticky=(tk.W + tk.E)
                          )

        LabelInput(post_default, '', input_class=ttk.Radiobutton,
                   var=self._vars['Post Job'],
                   input_args={'values': ['Delete', 'Keep']}
                   ).grid(row=1, column=0, sticky=(tk.W + tk.E)
                          )

        LabelInput(zip_default, '', input_class=ttk.Radiobutton,
                   var=self._vars['Zip?'],
                   input_args={'values': ['Yes', 'No']}
                   ).grid(row=2, column=0, sticky=(tk.W + tk.E)
                          )

        self.run_button = tk.Button(buttons, text='Run',
                                    command=self._on_run
                                    )
        self.run_button.grid(row=0, column=1, sticky='ew')

        self.select_target = tk.Button(buttons, text='Select Target',
                                       command=self._on_target
                                       )
        self.select_target.grid(row=0, column=0, sticky='ew')

        self.quit_button = tk.Button(
            buttons,
            text='Quit',
            command=self._on_quit
        )
        self.quit_button.grid(row=3, column=0, sticky='ew')

        self.status = tk.StringVar(
            None, 'Status: '
                  'Please select a target to continue...'
        )
        ttk.Label(
            self, textvariable=self.status, wraplength=225, justify='left'
        ).grid(sticky=(tk.W + tk.E), row=4, padx=10)

    def _on_run(self):
        """Command to run scrubber on target(s)"""

        # Check if target is valid and update status
        if self.input_folder == '' and self.input_file == '':
            if self._vars['Batch Size'].get() == 'Folder':
                self.status.set('No target chosen. \n'
                                'Please choose a target folder...')
            else:
                self.status.set('No target chosen. \n'
                                'Please choose a target zip file...')
        else:
            # Process zip file(s)
            if self.input_folder == '':
                self.status.set(
                    f'Processing {self.input_file} ...')
                self.update()
                # Process target zip
                self.err_present = process_exports(self.input_file,
                                                   self._vars['Post Job'].get(),
                                                   self._vars['Zip?'].get(),
                                                   )
                self.input_file = ''
            else:
                for file in os.listdir(self.input_folder):
                    self.status.set(f'Processing {file} ...')
                    self.update()
                    # Process all zips in target directory
                    if '.zip' in file:
                        self.err_present = process_exports(
                            self.input_folder + '/' + file,
                            self._vars['Post Job'].get(),
                            self._vars['Zip?'].get(),
                        )
                        if self.err_present == 1:
                            self.err_count += 1
                self.input_folder = ''

            # Check and alert on errors present during run
            if self.err_count == 0 and self.err_present == 0:
                self.status.set('Processing Complete. '
                                'Add more targets to continue.')
            else:
                self.status.set('Processing Complete, '
                                'but errors are present.'
                                '\nPlease refer to the error file which will '
                                'be contained in the target directory.')
                self.err_count = 0

    def _on_target(self):
        """Command to choose a target folder/file"""

        if self._vars['Batch Size'].get() == 'Folder':
            # Request target directory
            ch_folder_diag = tk.Tk()
            ch_folder_diag.overrideredirect(True)
            ch_folder_diag.attributes('-alpha', 0)
            ch_folder_diag.title('Choose target folder...')
            self.input_folder = filedialog.askdirectory(
                title='Choose target folder...')
            ch_folder_diag.destroy()
            self.input_file = ''
        else:
            # or request target zip file
            ch_file_diag = tk.Tk()
            ch_file_diag.overrideredirect(True)
            ch_file_diag.attributes('-alpha', 0)
            ch_file_diag.title('Choose target file...')
            self.input_file = filedialog.askopenfilename(
                title='Choose target file...',
                filetypes=[("Zip Files", "*.zip")])
            ch_file_diag.destroy()
            self.input_folder = ''

        # Update status label with target info
        if self.input_folder != '':
            self.status.set(
                f'Target folder set to: \n{self.input_folder}. '
                f'\nChoose Run to continue...'
            )
        else:
            self.status.set(f'Target file set to: \n{self.input_file}. '
                            f'\nChoose Run to continue...'
                            )

    @staticmethod
    def log_error(err_file, message) -> None:
        """Simple method for opening passed txt file
        and appending message

        :param err_file: txt file
        :param message: string to append.
        This method will prepend current date and time
        """

        with open(err_file, 'a') as f:
            f.write(f'Error_{datetime.datetime.now()}_{message}\n')

    @staticmethod
    def _on_quit():
        """Command to exit program"""
        sys.exit()
