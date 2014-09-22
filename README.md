police-stats-scraper
====================

The police stats are released in an unusable format. This scraper fixes that.

Usage: ./scrape.py <police stats.xls> <start year> <end year>

Currently the police stats are released per police station in a way that doesn't allow for pivoting or any sort of analysis. This scraper outputs the data in a saner format with the following headings:

Province,Police Station,Crime,Year,Incidents
