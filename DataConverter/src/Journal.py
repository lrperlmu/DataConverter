
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

# file: Journal.py
# author: Leah Perlmutter
# date: 7 June 2012

import csv
from DataConverter import *

# Helps make construction of JournalItems generic
class JournalItemTemplate:
    
    acct = -1
    desc = ""

    # relevantHeaders = [ts_dept, ts_sales, variance, sears_rent, over_short,
    # sales_tax, total_pay, sears_sales, commission]
    header = []
    sign = []

    # produce an item only if balance nonzero, e.g. cash over/short
    nonzeroCond = False

    ## e.g. Variance value is 
    ## + sears_sales - ts_sales
    ## header = [sears_sales, ts_sales]
    ## sign = [+1.0, -1.0]
    ##
    ## most will only have one element in the lists, e.g.
    ## TS sales
    ## header = [ts_sales]
    ## sign = [+1]
    

    # Two JournalItemTemplates are equal if their fields are equal
    def __eq__(self, other):
        if (self.acct != other.acct or 
            self.desc != other.desc or
            self.header != other.header or
            self.sign != other.sign):
            return False
        else:
            return True

    # print will print the fields of a JournalItemTemplate
    def __str__(self):
        ret =  "acct:" + self.acct.__str__() + \
            " desc:" + self.desc + \
            " headers:" + self.header.__str__() + \
            " signs:" + self.sign.__str__()
        return ret


# One credit or debit for one account
# also associated with department number and description
class JournalItem: 

    special = "" # "blank", "header", "total",  "", "custom"
    custom = [] # optional fields for special entries
    
    # fields for normal entries
    acct = -1 # account number
    value = 0 # positive = credit, negative = debit
    dept = -1 # department number
    desc = "" # description to write out

    # initialize all fields to given or default values
    def __init__(self, special="", acct=-1, \
                     value=0, dept=-1, desc="", custom=[]):
        self.special = special
        self.acct = acct
        self.value = value
        self.dept = dept
        self.desc = desc
        self.custom = custom

    # return list representation of this journal item
    # suitable for use with csv writer
    def toList(self):
        if self.special=="":
            if self.value<0:
                return [self.acct, self.dept, "%.02f" %(-self.value), \
                            None, self.desc]
            else:
                return [self.acct, self.dept, None, \
                            "%.02f" %(self.value), self.desc]

        elif self.special=="blank":
            return []
        
    # return string representation of this journal item
    # for debugging purposes (not escaped properly for csv file)
    def __str__(self):
        ret = str(self.acct) + "," + str(self.dept) + ","

        if self.value<0:
            ret += "%.02f" %(-self.value) + ",," # debit column
        else:
            ret += "," + "%.02f" %(self.value) + "," # credit column

        ret += self.desc + "," + "\n"

        return ret


# A list of journal items
class JournalEntry:

    # list of strings holds the meaning of each column
    header = []
    
    # list of JournalItem
    items = []

    # maps account/description keys to values
    totals = {}

    # optionally initialize header
    def __init__(self, header=[]):
        self.header = header
        self.totals["Balance"] = 0.
        

    # put the given journal item in the list of items
    # add its value to the relevant running totals
    def put(self, item):
        self.items.append(item)

        # for normal entries only, add to totals
        if item.special=="":

            # overall total
            self.totals["Balance"] += item.value

            # account/description total
            key = self.makeKey(item.acct, item.desc)
            if self.totals.has_key(key):
                self.totals[key] += item.value
            else:
                self.totals[key] = item.value

    # this is a method in case we want to change the way it's done
    def makeKey(self, number, string):
        return number.__str__() + " " + string

    # put an entry for each template in items
    # and one for the balance
    def putTotals(self, items):
        self.put(JournalItem(special="custom", 
                             custom=["Description", None, None, "Value"]))

        for template in items:

            # concatenate acct with desc -- that is the key in the dict
            key = self.makeKey(template.acct, template.desc)
            cells = ["Total " + key, None, None, writeFloat(self.totals[key])]
            item = JournalItem(special="total", custom=cells)
            self.put(item)

        cells = ["Balance", None, None, writeFloat(self.totals["Balance"])]
        item = JournalItem(special="total", custom=cells)
        self.put(item)

        self.put(JournalItem(special="blank"))

    # Write all items to given csv writer
    def writeCSV(self, filename):

        outfileobj = open(filename, 'wb')
        writer = csv.writer(outfileobj, dialect='excel')

        # write all the items
        for item in self.items:

            # normal item: 
            if item.special=="":
                writer.writerow(item.toList())

            # header
            elif item.special == "header":
                writer.writerow(self.header)

            # blank line
            elif item.special == "blank":
                writer.writerow([])

            # total
            elif item.special == "total":
                writer.writerow(item.custom)

            # custom
            elif item.special == "custom":
                writer.writerow(item.custom)

        outfileobj.flush()
        outfileobj.close()

    

