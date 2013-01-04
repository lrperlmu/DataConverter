
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

# file: DCExceptions.py
# author: Leah Perlmutter
# date: 12 June 2012

# Exceptions for Data Converter tool

# Data Converter Error
# base class for other exceptions here
class DCError(Exception):
    title = "Data Converter Error"
    message = "unknown error"


# To be raised if the headers of the input file are wrong
class HeaderError(DCError):
    title = "Header Error"

    # notfound - list of missing headers
    def __init__(self, notfound):
        self.message = "Missing header(s):\n\n"
        for header in notfound:
            self.message += "'" + header + "'\n"
        self.message += "\n" + \
            "Please check format of input file. Verify that all \n" + \
            "required headers are present and appear in the same row." 

    def __str__(self):
        return self.message

class FileTypeError(DCError):
    title = "File Type Error"

    def __init__(self):
        self.message = \
            "Please select a file that can be opened in Excel as a\n" + \
            'worksheet. Suitable file extensions include ".xls", ".xlsx",\n'+\
            'and ".csv".'
        
    def __str__(self):
        return self.message
    


# To be raised if the own/sub file is wrong
class DepartmentNumberError(DCError):    
    title = "Department Number Error"

    def __init__(self, storenum, linenum):
        self.message = \
            "Department number %d found in row %d of input file, but not \n" \
            %(storenum, linenum)+ \
            "in stores.csv. Please make sure that stores.csv is up to date." \

        
    def __str__(self):
        return self.message
 
