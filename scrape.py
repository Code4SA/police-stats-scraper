"""
A simple scraper the converts the crime states excel file released by SAPS into something more usable.

usage: ./scrape.py <police stats.xls> <start year> <end year>
"""

from __future__ import print_function
from __future__ import division
import sys
import re
import xlrd
import csv

class PoliceStateException(Exception):
    pass

class PoliceState(object):
    def __init__(self, rows):
        self.rows = rows
        self.curstate = self.start_state

        # e.g. Amangwe (KZN) for April to March 2004/2005 - 2013/2014
        self._re_header = re.compile("""
            ([\w\s]+\w)      # Station Name
            .* for .* to .* 
            (\d{4})\/(\d{4}) # End Period (i.e. 2013/2014)
            \s*$
        """, re.VERBOSE)

    
    def skip(self, num):
        self.rows = self.rows[num:]
        return self.rows

    def get_data(self):
        data = {
            "crimes" : {}
        }
        while self.curstate != None:
            data = self.curstate(data)
        return data

    def start_state(self, data):
        row = self.rows[0]
        if not "Crime Research and Statistics - South African police Service" in row[0]:
            raise PoliceStateException("Data not aligned correctly")

        
        (police_station, syear, eyear) = self._re_header.search(self.rows[2][0]).groups()

        self.skip(7)
        self.curstate = self.section_state
        
        data["name"] = police_station
        data["start_year"] = syear
        data["end_year"] = eyear

        
        data.update({
            "name" : police_station,
            "start_year" : syear,
            "end_year" : eyear,
        })

        return data

    def section_state(self, data):
        self.skip(1)
        try:
            while True:
                row = self.rows[0]
                crime = row[0]

                if row[1] == "":
                    return data

                data["crimes"][crime.strip()] = row[1:]
                self.skip(1)
        except IndexError:
            self.curstate = None
            return data
        

def process_sheet(worksheet):
    rows = [worksheet.row_values(i) for i in range(worksheet.nrows)]
    def align_with_first_section(rows):
        for idx, row in enumerate(rows):
            if "Crime Research and Statistics - South African police Service" in row[0]:
                rows = rows[idx:]
                return jump_section(rows)

    def jump_section(rows):
        return rows[45:]

    def grab_section(rows):
        return rows[0:45]

    rows = align_with_first_section(rows)
        
    while len(rows) > 0:
        extract = grab_section(rows)
        ps = PoliceState(extract)
        data = ps.get_data()
        data["province"] = worksheet.name
        rows = jump_section(rows)
        yield data

def write_csv(data, start_year, end_year):
    years = range(start_year, end_year + 1)
    writer = csv.writer(sys.stdout)
    name = data["name"]
    province = data["province"]

    crimes = sorted(data["crimes"].keys())
    for crime in crimes:
        for idx, year in enumerate(years):
            writer.writerow([province, name, crime, year, data["crimes"][crime][idx]])

def main(args):
    filename = args[1]
    start_year, end_year = int(args[2]), int(args[3])
    workbook = xlrd.open_workbook(filename)
    sheet_names = workbook.sheet_names()
    print("Province,Police Station,Crime,Year,Incidents")
    
    for sheet_name in sheet_names:
        worksheet = workbook.sheet_by_name(sheet_name)
        for data in process_sheet(worksheet):
            write_csv(data, start_year, end_year)

if __name__ == "__main__":
    main(sys.argv)
