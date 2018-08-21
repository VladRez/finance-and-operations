"""
Remove commas between quotes to avoid truncating strings
 when importing into SQL server table.
 Example "Apple, Inc" -> "Apple Inc"
"""

import re

db = open('C:/dataset/MU_REPORT.csv').readlines()
db_clean =  open('C:/dataset/MU_REPORT_CLEAN.csv', 'w')
rows = set()
for row in db:
    db_clean.write(re.sub(r'(?!(([^"]*"){2})*[^"]*$),', '', row))

db_clean.close()
