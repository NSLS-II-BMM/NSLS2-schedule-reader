# NSLS2-schedule-reader

Simple script to consume an NSLS2 schedule spreadsheet and produce ICS
calendar files.

To use:
```
usage: NSLS2xlsx2ical.py [-h] [--month MONTH] xlsx

Transmogrify the NSLS2 Excel schedule to ICS files.

positional arguments:
  xlsx           Excel file

optional arguments:
  -h, --help     show this help message and exit
  --month MONTH  Month,Year
```

For example:
```
./NSLS2xlsx2ical.py FY23.xlsx --month='September,2023'
```
will read a schedule in a file called `FY23.xlsx` output a file called
`NSLS2-September,2023-ops.ics` while
```
./NSLS2xlsx2ical.py FY23.xlsx
```
will present you with a list of possible months:
```
  1:   September, 2022
  2:   October, 2022
  3:   November, 2022
  4:   December, 2022
  5:   January, 2023
  6:   February, 2023
  7:   March, 2023
  8:   April, 2023
  9:   May, 2023
 10:   June, 2023
 11:   July, 2023
 12:   August, 2023
 13:   September, 2023

Choose a month [1-13, q=quit] 
```

## ICS files for FY23 (September 2022 - September 2023)

* [September, 2022](./FY23/NSLS2-September,2022-ops.ics)
* [October, 2022](./FY23/NSLS2-October,2022-ops.ics)
* [November, 2022](./FY23/NSLS2-November,2022-ops.ics)
* [December, 2022](./FY23/NSLS2-December,2022-ops.ics)
* [January, 2023](./FY23/NSLS2-January,2023-ops.ics)
* [February, 2023](./FY23/NSLS2-February,2023-ops.ics)
* [March, 2023](./FY23/NSLS2-March,2023-ops.ics)
* [April, 2023](./FY23/NSLS2-April,2023-ops.ics)
* [May, 2023](./FY23/NSLS2-May,2023-ops.ics)
* [June, 2023](./FY23/NSLS2-June,2023-ops.ics)
* [July, 2023](./FY23/NSLS2-July,2023-ops.ics)
* [August, 2023](./FY23/NSLS2-August,2023-ops.ics)
* [September, 2023](./FY23/NSLS2-September,2023-ops.ics)
