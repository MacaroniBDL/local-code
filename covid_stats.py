import os
#methods to better index values
import string
import time
from datetime import date

import openpyxl
from openpyxl import load_workbook


def col2num(col):
    num = 0
    for c in col:
        if c in string.ascii_letters:
            num = num * 26 + (ord(c.upper()) - ord('A')) + 1
    return num

def num2col(n):
    """Number to Excel-style column name, e.g., 1 = A, 26 = Z, 27 = AA, 703 = AAA."""
    name = ''
    while n > 0:
        n, r = divmod (n - 1, 26)
        name = chr(r + ord('A')) + name
    return name


fdir = input("please paste local A Teams path: ")
while not os.path.exists(fdir):
    print("That was an invalid path please try again")
    fdir = input("Enter a valid Path: ")



def main(fdir):
    ftemp = f"{fdir}\\NYT Case Avg. Yesterday.xlsx"

    wb = load_workbook(filename=ftemp)

    #get data from sheet us-counties-all-latest-avg
    #put in B-Format
    b_form = wb["B-Format"]
    county_data = wb["us-counties-all-latest-avg"]
    #first time run only then comment out
    #append the full name for the state in column b of B-Format

    state_abbr = wb["state-abbreviations"]


    #compare current cell against column B of state_abbr, result will be in column A
    #---------------------------------------------------------------------------------
    row_b_form = 2
    col_b_form = 1
    c_cell = b_form[num2col(col_b_form) + str(row_b_form)].value
    while c_cell != None:
        sa_row = 1
        abbr_cell = state_abbr['B' + str(sa_row)].value

        while abbr_cell != None and abbr_cell != c_cell:
            sa_row += 1
            abbr_cell = state_abbr['B' + str(sa_row)].value

        if abbr_cell == None:
            print("State abbr. NOT FOUND", c_cell, row_b_form)
        else:
            state_name = state_abbr['A' + str(sa_row)].value
            b_form['B' + str(row_b_form)].value = state_name
        row_b_form += 1

        c_cell = b_form['A' + str(row_b_form)].value
    #-------------------------------------------------------------------------------


    #clean input string before searching
    #combine column B and C, B-Format and C B for state_abbr
    #add column header in nytcaseavg to display date
    row_b_form = 1
    col_b_form = 1
    c_cell = b_form[num2col(col_b_form) + str(row_b_form)].value
    #method to find what column to add the date in
    while c_cell != None:
        col_b_form += 1
        c_cell = b_form[num2col(col_b_form) + str(row_b_form)].value

    today = date.today()
    d_formatted = today.strftime("%m/%d/%Y")

    b_form[num2col(col_b_form) + str(1)].value = d_formatted

    row_b_form = 2
    c_cell = b_form['C' + str(row_b_form)].value

    #now comparing b-format to state abbr.
    while c_cell != None:
        in_str = (b_form['B' + str(row_b_form)].value + c_cell).upper()
        in_str = in_str.replace(' ', '')
        row_counties = 2
        county = county_data['B' + str(row_counties)].value
        out_str = (county_data['C' + str(row_counties)].value + county).upper()
        out_str = out_str.replace(' ', '')
        while county != None and out_str != in_str:
            row_counties += 1
            county = county_data['B' + str(row_counties)].value
            if county != None:
                out_str = (
                    county_data['C' + str(row_counties)].value + county).upper()
                out_str = out_str.replace(' ', '')

        if county == None:
            print("county not found: ", in_str)
        else:

            b_form[num2col(col_b_form) + str(row_b_form)
                ].value = county_data['D' + str(row_counties)].value

        row_b_form += 1
        c_cell = b_form['C' + str(row_b_form)].value


    d = today.strftime("%d-%m")
    fname = f"{fdir}\\B-FORMAT {d}.xlsx"
    wb.save(fname)
    print("file saved")
    

try:
    print("Please wait while I format your data")
    main(fdir)
    print("okay done here!")
    print("\nYour file will placed in A-Teams under the name B-Format dd-mm.\n For further assistance please contact an IT member")
except Exception as e:
    print("something bad happened contact IT department")
    ans = input("display error?")
    if ans == 'y':
        print(e)


time.sleep(5)

