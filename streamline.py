import io
import os
from datetime import date, timedelta

import openpyxl as opyxl
import pandas as pd
import requests
from openpyxl.utils.dataframe import dataframe_to_rows
from pandasql import sqldf

states_url = "https://github.com/nytimes/covid-19-data/raw/master/rolling-averages/us-states.csv"
states_download = requests.get(states_url).content

df_states = pd.read_csv(io.StringIO(states_download.decode('utf-8')))

counties_url = "https://raw.githubusercontent.com/nytimes/covid-19-data/master/rolling-averages/us-counties-recent.csv"
counties_download = requests.get(counties_url).content

df_counties = pd.read_csv(io.StringIO(counties_download.decode('utf-8')))

#after creating data frames fro each of the 
#print(df_states.head)
yday = date.today() - timedelta(days = 1)
yday = str(yday)

q_states = f"SELECT date, state, cases_avg FROM df_states WHERE date = '{yday}'"
q_counties = f"SELECT date, county, state, cases_avg FROM df_counties WHERE date = '{yday}'"
yday_state_results = sqldf(q_states)
yday_counties_results = sqldf(q_counties)

#------------------------------------------------------------------------------------------------------------------------------
#comments for my sanity
# in order to make this into a valuable script I will make everything into functions that can be called through prompts,
# I can even allow writing SQLite statements
# after the data has been collected it must be put into a workbook
# making a copy of nyt covid....
#------------------------------------------------------------------------------------------------------------------------------

fname = 'NYT Case Avg. Yesterday.xlsx'


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



def save_data(fname):
    wb = opyxl.load_workbook(filename=fname)
    b_form = wb["B-Format"]
    county_data = wb["us-counties-all-latest-avg"]
    state_data = wb["us-states-latest-avg"]
    state_abbr = wb["state-abbreviations"]
    #states dataframe to excel
    for r in dataframe_to_rows(yday_state_results, index = False, header = True):
        state_data.append(r)
    #counties dataframe to excel
    for r in dataframe_to_rows(yday_counties_results, index = False, header = True):
        county_data.append(r)

    wb.save(fname)    

def main(fdir):
    ftemp = f"NYT Case Avg. Yesterday.xlsx"

    wb = opyxl.load_workbook(filename=ftemp)

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
                ].value = int(county_data['D' + str(row_counties)].value)

        row_b_form += 1
        c_cell = b_form['C' + str(row_b_form)].value


    d = today.strftime("%d-%m")
    fname = f"B-FORMAT {d}.xlsx"
    wb.save(ftemp)
    print("file saved")

fdir = os.getcwd
save_data(fname)
main(fdir)