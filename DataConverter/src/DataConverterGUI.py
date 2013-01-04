
''' 
This file is part of Data Converter.

Data Converter is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by 
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful, 
but without any warranty; without even the implied warranty of 
merchantability or fitness for a particular purpose. See the 
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with Data Covnerter. If not, see <http://www.gnu.org/licenses/>.
'''


# file: DataConverterGUI.py
# author: Leah Perlmutter
# date: 7 June 2012

import DataConverter

import sys
import os
from Tkinter import Tk
import tkFileDialog 
import tkMessageBox 

class DataConverterGUI:

    # Display a welcome dialog box and ask the user to continue
    def __init__(self):
        root = Tk()

        # use the python icon on all windows
        iconpath = "\\Python27\\DLLs\\py.ico" ## magic constant, change
        root.tk.call('wm', 'iconbitmap', root._w, '-default', iconpath)

        root.withdraw() # don't show root window

        msg = \
            "This application reads the Sears monthly spreadsheet\n" + \
            "and writes relevant data to an excel file.\n\n\n" + \
            "Data Converter Copyright 2012 Leah Perlmutter.\n" + \
            "Distributed with no warranty under the GNU General\n" +\
            "Public License, which can be viewed at \n" + \
            "<http://www.gnu.org/licenses/>. \n" + \
            "Licensees may convey the work under this license.\n\n\n" + \
            'Click "OK" to open the Sears spreadsheet.'
        con = tkMessageBox.askokcancel(title="Welcome to Data Converter",
                              icon="info",
                              message=msg)
        if not con:
            sys.exit()

    # Report to the user that we are finished and offer
    # to open the output file in excel.
    def report(self, outfile):
        # dialog box: Report completion and ask to open output file
        opn=tkMessageBox.askyesno(title='Data Collected Successfully', \
                               message='Output file saved as ' + outfile + \
                               '.\n\nWould you like to open the file in ' + \
                               'Excel now?')
        if opn:
            # run a vbs script to open it and format columns
            command = 'openInExcel.vbs "' + outfile + '"'
            os.system(command)

    # make forward slashes into backslashes
    # "user/Desktop/dir" --> "user\Desktop\dir"
    def backSlash(self, filename):
        return filename.replace("/", "\\")

    # Give the user a file picker dialog box
    # Return path of file picked or default if none picked
    def promptOpenFile(self):
        name = tkFileDialog.askopenfilename()
        name = self.backSlash(name)
        return name

    # Display an exception's message in a dialog box 
    def handleExceptionRestart(self, ex):

        message = ex.message + \
            '\n\nClick "OK" to return to Data Converter\'s Welcome screen.'
        tkMessageBox.showerror(title=ex.title, icon="error", message=message)
        
