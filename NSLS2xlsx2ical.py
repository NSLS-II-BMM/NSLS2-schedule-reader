#!/usr/bin/env python3

import argparse
from NSLS2Calendar import NSLS2Calendar

parser = argparse.ArgumentParser(description='Transmogrify the NSLS2 Excel schedule to ICS files.')
parser.add_argument('xlsx', metavar='xlsx', type=str, help='Excel file')
parser.add_argument('--month', type=str, help='Month,Year', required=False)
args = parser.parse_args()

nsls2calendar = NSLS2Calendar()
nsls2calendar.set_workbook(args.xlsx)
nsls2calendar.current_month(args.month)
nsls2calendar.write_ics()
