#!\Python27\python.exe


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

# file: DataConverter.py
# author: Leah Perlmutter
# date: 1 June 2012

# Developed for:
# Windows XP Pro
# Excel 2007 (12.0.6661.5000)
# Python 2.7.3

# input: csv file -- Barb's version of the Sears monthly report
# output: csv file -- journal entry for the AS400

from Journal import *
from DataConverterGUI import *
from DCExceptions import *

import csv
import os
import copy
import sys
import datetime

# Debug printing
global printf
#printf = sys.stdout.write
def donothing(s):
    pass
printf = donothing


# return a reader for the csv file we want to read
def getCSV(csvpath):

    # default file
    localname = "Sears - October, 2012.xls"
    excelfile = os.getcwd() + "\\..\\io\\" +  localname 

    # prompt user for file
    if GUI:
        excelfile = GUI.promptOpenFile()
        if excelfile == "":
            main(sys.argv)
            return
    else:
        prompt = "Input file (default " + ".\\" + localname + "):"
        answer = raw_input(prompt)
        if len(answer)>0:
            excelfile = os.getcwd() + "\\" + answer

    # Check that user selected file that excel can open
    if not (excelfile.endswith(".csv") or 
            excelfile.endswith(".xls") or
            excelfile.endswith(".xlsb") or
            excelfile.endswith(".xml") or
            excelfile.endswith(".xlw") or
            excelfile.endswith(".xlsx") ):
        raise FileTypeError()

    # Convert the .xls file to a temporary .csv file so we can read it
    print "deleting the old tmp file in case it was left over..."
    os.system('erase ' + '"' + csvpath + '"') # in case old one is left over
    command = 'xlsToCsv.vbs "' + excelfile + '" "' + csvpath + '"'
    os.system(command)
    
    # return a csv reader to the csv file
    fileobj = open(csvpath, 'rb')
    reader = csv.reader(fileobj, dialect='excel')
    return fileobj, reader


# Takes in a list of lines of the csv file
# each line is a list of entries
# Writes the header dictionary mapping column meanings to column numbers
def getHeaders(inputlines):

    ##
    global ts_dept, ts_sales, variance, sears_rent, \
        over_short, sales_tax, total_pay, sears_sales, commission

    # header string constants
    ts_dept = "ts dept"
    ts_sales = "ts sales"
    variance = "variance"
    sears_rent = "sears rent"
    over_short = "over / short"
    sales_tax = "net sales tax"
    total_pay = "total pay"

    sears_sales = "net concession sales"
    commission = "ts com $"

    relevantHeaders = [ts_dept, ts_sales, variance, sears_rent, over_short,
                       sales_tax, total_pay, sears_sales, commission]

    # maps header id strings to their column number
    global headerdict
    headerdict = {}

    # find the line with the headers 
    #  (assume to be first line with first cell fillled)
    ## (error message if header line not found?)
    i = 0
    for line in inputlines:
        if len(line) > 0 and len(line[0]) > 0:
            headerline = line
            break
        i = i+1

    #print headerline

    # parse the header line into a dict
    n = 0
    for word in headerline:
        lword = word.lower() # case insensitive
        for header in relevantHeaders:
            if lword.find(header) > -1: # found
                headerdict[header] = n
                relevantHeaders.remove(header)
        n = n+1

    if len(relevantHeaders) > 0:
        raise HeaderError(relevantHeaders)



# parse the own/sub file and build a dictionary
def buildOwnSubDict():
    ownsubfile = os.getcwd() + "\\..\\config\\stores.csv"

    # make a csv reader to read the own sub file
    fileobj = open(ownsubfile, "rb")
    reader = csv.reader(fileobj, dialect="excel")

    # maps store numbers to "own" or "sub"
    ownsubdict = {}
    
    for line in reader:
        if len(line) > 1: # line has at least 2 entries
            word = line[0].strip().lower()
            if word == "sub" or word == "own":
                storenum = int(line[1]) 
                ownsubdict[storenum] = word

    return ownsubdict


# Parse a numeric string in accounting language.
# Imitates behavior of float() and interprets numbers in 
# parentheses as negatives
# "(100)" --> -100.0
# "-100"  --> -100.0
#  "100"  -->  100.0
# "(-100)" --> 100.0 (though we don't expect to see this case)
#
## should throw useful exception when it gets the empty string.
## could be because there was an unexpected empty cell in the
## spreadsheet. --> wrong own/sub dictionary?
def parseFloat(valstring):
    tmp = valstring.strip() # strip spaces
    if tmp.startswith("(") and tmp.endswith(")"):
        tmp = tmp.strip("()")
        return -float(tmp)
    else:
        return float(tmp)

# Write a float in accounting language.
# Encloses negatives in parentheses.
# -100 --> "(100.00)"
#  100 -->  "100.00"
def writeFloat(value):
    if value >= 0:
        return "%0.2f" %(value)
    else:
        return "(%0.2f)" %(-value)


# params:
#    journal - the journal entry to add items to
#    inputLine - (list of strings) - a row from the input file
#    storenum - store number for the output items we're putting
#    items - list of JournalItemTemplate - put a JournalItem for each

def doOutput(journal, inputLine, storenum, items):
    
    journal.put(JournalItem(special="header"))

    for template in items:
        entry = JournalItem()
        entry.dept = storenum
        entry.acct = template.acct
        entry.desc = template.desc

        entry.value = 0.0
        for (header, sign) in map(None, template.header, template.sign):
            entry.value += parseFloat(inputLine[headerdict[header]]) * sign

        #if (template.nonzeroCond==False) or (entry.value!=0.0):
        journal.put(entry)
        printf(entry.__str__())

    journal.put(JournalItem(special="blank"))


# An item is a value in the spreadsheet that we care about for each store.
#   e.g. total sales, sales tax
# The template for that item stores which account it goes into, whether
#   that value gets added or subtracted to the total in the account,
#   and a text description of that item. The text description is used
#   for finding the item in the spreadsheet.
# Own items are for owned stores
# Sub items are for subcontracted stores
# This method populates the lists of templates, one list for owned stores
#    and another list for subcontracted stores.
def outputTemplates():
    
    # own output
    ownItems = []

    template = JournalItemTemplate()
    template.acct = 12100
    template.desc = "TS Sales"
    template.header = [ts_sales]
    template.sign = [+1.0]
    ownItems.append(template)

    template = JournalItemTemplate()
    template.acct = 34700
    template.desc = "Variance"
    template.header = [sears_sales, ts_sales]
    template.sign = [+1.0, -1.0] 
    ownItems.append(template)

    template = JournalItemTemplate()
    template.acct = 54500
    template.desc = "Sears Rent"
    template.header = [sears_rent]
    template.sign = [-1.0] 
    ownItems.append(template)

    template = JournalItemTemplate()
    template.acct = 24000
    template.desc = "Sales Tax"
    template.header = [sales_tax]
    template.sign = [-1.0] 
    ownItems.append(template)

    template = JournalItemTemplate()
    template.acct = 18800
    template.desc = "Total Pay Received"
    template.header = [total_pay]
    template.sign = [-1.0] 
    ownItems.append(template)

    template = JournalItemTemplate()
    template.acct = 80500
    template.desc = "Cash Over / Short"
    template.header = [over_short]
    template.sign = [-1.0] 
    template.nonzeroCond = True
    ownItems.append(template)


    # sub output
    subItems = []
    
    template = JournalItemTemplate()
    template.acct = 11700
    template.desc = "Sales"
    template.header = [sears_sales]
    template.sign = [+1.0] 
    subItems.append(template)

    template = JournalItemTemplate()
    template.acct = 11700
    template.desc = "Rent"
    template.header = [sears_rent]
    template.sign = [-1.0] 
    subItems.append(template)

    template = JournalItemTemplate()
    template.acct = 11700
    template.desc = "Sales Tax"
    template.header = [sales_tax]
    template.sign = [-1.0] 
    subItems.append(template)

    template = JournalItemTemplate()
    template.acct = 18800
    template.desc = "Total Pay Received"
    template.header = [total_pay]
    template.sign = [-1.0] 
    subItems.append(template)

    template = JournalItemTemplate()
    template.acct = 11700
    template.desc = "Cash Over / Short"
    template.header = [over_short]
    template.sign = [-1.0] 
    template.nonzeroCond = True
    subItems.append(template)

    template = JournalItemTemplate()
    template.acct = 11700
    template.desc = "Commission"
    template.header = [commission]
    template.sign = [-1.0] 
    subItems.append(template)

    template = JournalItemTemplate()
    template.acct = 82500
    template.desc = "Commission"
    template.header = [commission]
    template.sign = [+1.0] 
    subItems.append(template)

    template = JournalItemTemplate()
    template.acct = 39100
    template.desc = "Sales"
    template.header = [sears_sales]
    template.sign = [+1.0] 
    subItems.append(template)

    template = JournalItemTemplate()
    template.acct = 12000
    template.desc = "Sales Offset"
    template.header = [sears_sales]
    template.sign = [-1.0] 
    subItems.append(template)

    template = JournalItemTemplate()
    template.acct = 54500
    template.desc = "Rent"
    template.header = [sears_rent]
    template.sign = [-1.0] 
    subItems.append(template)

    template = JournalItemTemplate()
    template.acct = 82700
    template.desc = "Sub Rent"
    template.header = [sears_rent]
    template.sign = [1.0] 
    subItems.append(template)

    return ownItems, subItems


def handleExceptionRestart(ex):
    if GUI:
        GUI.handleExceptionRestart(ex)
    else:
        print ex.message
        print "\n\nTry again."
    main(sys.argv)
        

# Read in the Sears monthly report and write out a csv file
# with data to enter into the AS400
def main(argv):

    # use gui or not?
    global GUI
    if len(sys.argv)>1 and sys.argv[1] == "-nogui":
        GUI = None
    else:
        GUI = DataConverterGUI()
        
    # read the csv file
    tmppath = os.getcwd() + "\\.tmpSearsIn.csv"
    try:
        (fileobj, reader) = getCSV(tmppath)
    except DCError as ex:
        handleExceptionRestart(ex)
        return
    inputlines = []
    for line in reader:
        inputlines.append(line)
    fileobj.close()

    try:
        getHeaders(inputlines) # assigns headerdict
    except DCError as ex:
        handleExceptionRestart(ex)
        return

    # read in the own/sub key file
    ownsubdict = buildOwnSubDict() # assigns ownsubdict

    # init journal
    journal = JournalEntry(["Acct #", "Dept #", "Debit", "Credit", "Description"])

    # make item templates
    (ownItems, subItems) = outputTemplates()

    # go over lines of file
    storeidx = headerdict[ts_dept] # which column of input has storenums?
    linenum = 1
    for line in inputlines:

        # skip any line that doesn't have a number in the store-number column
        if (len(line)<storeidx) or (not line[storeidx].strip().isdigit()):
            linenum += 1
            continue
        
        # see if it's an own store or a sub and do the right thing
        storenum = int(line[storeidx])
        if not storenum in ownsubdict:
            handleExceptionRestart(DepartmentNumberError(storenum, linenum))
        if ownsubdict[storenum] == "sub":
            doOutput(journal, line, storenum, subItems)
        elif ownsubdict[storenum] == "own":
            doOutput(journal, line, storenum, ownItems)

        linenum += 1

    # do totals, one for each item
    # Totals are for owned and subcontracted stores together,
    #   so we need to combine the lists of templates.
    allItems = []
    allItems.extend(ownItems)
    for item in subItems:
        if item not in allItems:
            allItems.append(item)

    journal.putTotals(allItems)

    # output file name
    time = datetime.datetime.now()
    stamp = time.__str__()
    stamp = stamp.split(".")[0].translate(None, ":").replace(" ", "-")
    outfile = os.getcwd() + "\\..\\io\\SearsOutput" + stamp + ".csv"

    # write output file
    journal.writeCSV(outfile)

    if GUI:
        GUI.report(outfile)
    else:
        print "Output saved as " + outfile + "."
    
    # finally?
    # remove .tmpSearsIn.csv when done
    os.system('erase "' + tmppath + '"' )


if __name__ == "__main__":
    main(sys.argv)

