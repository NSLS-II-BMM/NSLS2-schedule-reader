#!/usr/bin/env python3

from openpyxl import load_workbook
import datetime, sys, argparse
from dateutil import tz

class NSLS2Calendar():
    '''A class for consuming the NSLS2 schedule spreadsheet and producing
    ICS calendar files.

    nsls2calendar = NSLS2Calendar()
    nsls2calendar.set_workbook('schedule.xlsx')
    nsls2calendar.current_month('February,2023')
    nsls2calendar.write_ics() # --> "NSLS2-February,2023-ops.ics"

    '''
    wb = None 
    sheet = None 
    month = ''
    col = ''
    allmonths = []
    dates = []
    startcol = []
    icol = 0
    calendar = []
    firstofmonth = None
    icsfile = ''
    

    def set_workbook(self, xlsx):
        self.wb = load_workbook(filename = xlsx)
        self.sheet = nsls2calendar.wb.worksheets[0]
        self.find_months()
    
    ## note: A to Z are characters 65 through 90
    def num2lett(self, i):
        '''Count columns in spreadsheet, i.e. A-Z then AA-AZ then BA-BZ and so on'''
        ii = int(i/26)
        if ii == 0:
            return '%s' % chr(64+i % 26)
        elif i/26 == 1 and i%26 == 0:
            return '%s' % chr(90)
        elif i%26 == 0:
            return '%s%s' % (chr(63+int(i/26)), chr(90))
        else:
            return '%s%s' % (chr(64+int(i/26)), chr(64+i % 26))

    def find_months(self):
        '''Find the columns in the spreadsheet that start each month

        The for loop goes out to 200.  This is a crude solution.

        A full year goes out to ~FX ~= 7*26 = 182, so 200 should be enough columns

        It's a bit hard to reliably find the right-most edge of the
        calendar since there are empty columns in between months.
        There is probably a solution to this problem...

        Note that the blue cells are in row 02

        '''
        for i in range(1,200):
            col = self.num2lett(i)
            ## find all the blue cells at the top with datetime-s in them
            try:
                #print(col, self.sheet[f'{col}02'].value, str(type(self.sheet[f'{col}02'].value)))
                if 'datetime.datetime' in str(type(self.sheet[f'{col}02'].value)):
                    self.allmonths.append(self.sheet['%s02' % col].value.strftime('%B,%Y'))
                    self.dates.append(self.sheet['%s02' % col].value)
                    self.startcol.append(i)
            except:
                pass

    def current_month(self, month=None):
        '''either take the month from the command line --month argument
        or present a dialog to the user to select a month

        Stuff the selected month's events into the calendar attribute
        '''
        if month is not None:
            if month in self.allmonths:
                self.firstofmonth = self.dates[self.allmonths.index(month)]
                self.icol = self.startcol[self.allmonths.index(month)]

        else:
            for (n,el) in enumerate(self.allmonths):
                print(f"{n+1:3}:   {el.replace(',', ', ')}")
            action = input(f"\nChoose a month [1-{len(self.allmonths)}, q=quit] ")
            if str(action) == 'q':
                print ('quitting...')
                exit()
            try:
                action = int(action)
            except:
                print(f'{str(action)} is not one of the choices.')
                exit()
            if action not in range(1, len(self.allmonths)+1):
                print(f'{str(action)} is not one of the choices.' )
                exit()
            month = self.allmonths[action-1]
            self.firstofmonth = self.dates[action-1]
            self.icol = self.startcol[action-1]

        self.month = month
        self.firstofmonth = self.firstofmonth.replace(tzinfo=tz.tzlocal())
        #print(self.allmonths, month)
        #print(month, self.firstofmonth)

        ## ----------------------------------------------------------------------
        ## now gather up the six columns that constitute the calendar for that month
        ## they are the six columns /after/ the column containing the name of the month
        halfshifts = list()
        for j in range(self.icol+1, self.icol+7):
            halfshifts.append(self.num2lett(j))

        current = self.firstofmonth

        ## ----------------------------------------------------------------------
        ## now read each half shift of the month, save a tuple with it's op type and a datetime
        for row in range(5, 36):
            for hs in halfshifts:
                if self.sheet[f'{hs}{row:02}'].value is not None:
                    current = current + datetime.timedelta(hours=4)
                    current = current.replace(tzinfo=tz.tzlocal())
                    self.calendar.append((self.sheet[f'{hs}{row:02}'].value, current))
                    #'%s%2.2d' % (hs, row)
                    
        ## would be nice to have a configured ics file naming pattern
        self.icsfile = f'NSLS2-{self.month}-ops.ics'

        # import pprint
        # pp = pprint.PrettyPrinter(indent=4)
        # pp.pprint(calendar)


    def write_ics(self):
        '''constructing an ICS object for the selected month
        '''
        from ics import Calendar, Event
        c = Calendar()

        begin = self.firstofmonth
        end   = None
        event = self.calendar[0][0]

        event_types = {'O':   'Accelerator Operations',
                       'D':   'Shutdown',
                       'S':   'Studies',
                       'I':   'Interlock Certification',
                       'M':   'Maintenance',
                       'S/M': 'Studies then Maintenance',
                       'M/S': 'Maintenance then Studies',
                       'O/M': 'Operations then Maintenance',
                       'S/O': 'Studies then Operations',
                       'C':   'Commissioning'}

        ## break the calendar into blocks of events, noting the begin
        ## and end times of each block
        for block in self.calendar:
            if block[0] != event:
                e = Event()
                e.name = event_types[event]
                e.begin = begin + datetime.timedelta(hours=-4)
                e.end = block[1] + datetime.timedelta(hours=-4)
                if event != 'O': c.events.add(e)
                event = block[0]
                begin = block[1]
        
        e = Event()
        e.name = event_types[event]
        e.begin = begin + datetime.timedelta(hours=-4)
        e.end = block[1] + datetime.timedelta(hours=-4)
        if event != 'O': c.events.add(e)

        with open(self.icsfile, 'w') as f:
            f.writelines(c)
        print(f'\nWrote {self.icsfile}')


parser = argparse.ArgumentParser(description='Transmogrify the NSLS2 Excel schedule to ICS files.')
parser.add_argument('xlsx', metavar='xlsx', type=str, help='Excel file')
parser.add_argument('--month', type=str, help='Month,Year', required=False)
args = parser.parse_args()

nsls2calendar = NSLS2Calendar()
nsls2calendar.set_workbook(args.xlsx)
nsls2calendar.current_month(args.month)
nsls2calendar.write_ics()
