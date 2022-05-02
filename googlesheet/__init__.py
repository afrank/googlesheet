import pickle
import os
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import sys
import logging
import more_itertools as mit

class GoogleSheet:
    def __init__(self, spreadsheet, sheet='Sheet1', cred_path="~/.config/google_sheet"):

        self.SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
        self.spreadsheet = spreadsheet

        self._sheet = sheet

        self.cred_path = os.path.expanduser(cred_path)
        
        self.auth = None
        self.last_cell = 0
        token_file = self.cred_path + '/token.pickle'
        google_creds_file = self.cred_path + '/credentials.json'

        if os.path.exists(token_file):
            with open(token_file, 'rb') as token:
                self.auth = pickle.load(token)
        
        if not self.auth or not self.auth.valid:
            if self.auth and self.auth.expired and self.auth.refresh_token:
                self.auth.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(google_creds_file, self.SCOPES)
                self.auth = flow.run_local_server(port=0)
            with open(token_file, 'wb') as token:
                pickle.dump(self.auth, token)
        
        self.service = build('sheets', 'v4', credentials=self.auth)

        self.all_cols = []
        for x in ["", "A","B","C"]:
            for y in range(26):
                v = x + chr(ord("A")+y)
                self.all_cols += [ v ]

        self.col = "A"
        self.row = 1

        self._src = self.read_sheet()

    @property
    def row(self):
        return self._row

    @property
    def col(self):
        return self._col

    @property
    def pos(self):
        return f"{self.col}{self.row}"

    @row.setter
    def row(self, row):
        self._row = row

    @col.setter
    def col(self, col):
        self._col = col

    @pos.setter
    def pos(self, pos):
        self.col = ''.join(filter(str.isalpha, pos))
        self.row = ''.join(filter(str.isdigit, pos))

    @property
    def sheet(self):
        return self._sheet

    @sheet.setter
    def sheet(self, sheet):
        self._sheet = sheet
        self._src = self.read_sheet()

    @property
    def x_col(self):
        return self.all_cols.index(self.col)

    @property
    def src(self):
        return self._src

    def col_x(self,x):
        return self.all_cols[int(x)]

    def set_pos(self,col,row):
        self.col = col
        self.row = row

    def __read(self,rnge):
        #logging.debug("Range: %s",rnge)
        result = self.service.spreadsheets().values().get(spreadsheetId=self.spreadsheet, range=rnge).execute()
        #logging.info('%s cells updated.',result.get('updatedCells'))
        values = result.get('values', [])
        return values

    def __write(self,body,rnge):
        logging.debug("Range: %s",rnge)
        result = self.service.spreadsheets().values().update(spreadsheetId=self.spreadsheet, range=rnge,valueInputOption='USER_ENTERED', body=body).execute()
        #print(body)
        #result = { "updatedCells": "Zarro" }
        logging.info('%s cells updated.',result.get('updatedCells'))

    @staticmethod
    def __fmt(val):
        if type(val) != list:
            return f"{val}"
        if type(val[0]) != list:
            return [ f"{x}" for x in val ]
        
        new = []
        for x in val:
            new += [ [ f"{y}" for y in x ] ]
        
        return new

    def write(self, x):
        """
        take val of any variety and flight control
        """
        #print("WRITE was called")
        #print(type(x))
        if type(x) != list:
            return self.write_one(x)

        if type(x[0]) != list:
            return self.write_vertical(x)

        return self.write_range(x)

    def write_one(self, val):
        """
        write a single cell
        takes one value
        """
        print("write_one called")
        body = { 'values': [[ self.__fmt(val) ]] }
        _range = f"{self.sheet}!{self.col}{self.row}:{self.col}{self.row}"
        self.__write(body,_range)

    def write_horizontal(self, vals):
        """
        write within a row across multiple columns
        takes one list of values
        """

        body = { 'values': [ self.__fmt(vals) ] }
        new_col = self.col_x(self.x_col + len(vals))
        _range = f"{self.sheet}!{self.col}{self.row}:{new_col}{self.row}"
        self.__write(body,_range)

    def write_vertical(self, vals):
        """
        write within multiple rows in a single column
        takes one list of values
        """
        print("write_vertical called")
        vals = [ [self.__fmt(x)] for x in vals ]
        body = { 'values': vals }
        new_row = int(self.row) + len(vals)
        _range = f"{self.sheet}!{self.col}{self.row}:{self.col}{new_row}"
        self.__write(body,_range)

    def write_range(self, vals):
        """
        write multiple rows across multiple columns
        takes a list of lists
        """
        print("write_range called")
        body = { 'values': self.__fmt(vals) }
        new_row = int(self.row) + len(vals)
        new_col = self.col_x(self.x_col + len(vals[0]))
        _range = f"{self.sheet}!{self.col}{self.row}:{new_col}{new_row}"
        self.__write(body,_range)

    def read(self):
        _range = f"{self.sheet}!{self.col}{self.row}:{self.col}{self.row}"
        return self.__read(_range)[0][0]

    def read_horizontal(self, cols=1):
        new_col = self.col_x(self.x_col + cols)
        _range = f"{self.sheet}!{self.col}{self.row}:{new_col}{self.row}"
        return self.__read(_range)

    def read_vertical(self, rows=1):
        new_row = self.row + rows
        _range = f"{self.sheet}!{self.col}{self.row}:{self.col}{new_row}"
        return self.__read(_range)

    def read_range(self, cols=1, rows=1):
        new_col = self.col_x(self.x_col + cols)
        new_row = self.row + rows
        _range = f"{self.sheet}!{self.col}{self.row}:{new_col}{new_row}"
        return self.__read(_range)

    def read_sheet(self, header=False, dict_key=""):
        """
        if header is true uses the first row as a header
        to create a list of dicts.

        if dict_key is supplied, it will return a dict.
        """

        sheet = self.read_range(len(self.all_cols)-1,500)

        if dict_key:
            rows = {}
        else:
            rows = []

        header = []
    
        for line in portfolio:
            if not line or line[0] == "":
                continue
            if not header:
                header = line
                continue
            row = dict(zip(header, line))
            if dict_key:
                k = row.get(dict_key)
                if k:
                    rows[k] = row
            else:
                rows += [ row ]
        
        return rows

    def down(self, cells=1):
        self.row += cells

    def up(self, cells=1):
        self.row = max(1, self.row - cells)

    def left(self, cells=1):
        self.col = self.col_x(self.x_col - cells)

    def right(self, cells=1):
        self.col = self.col_x(self.x_col + cells)

class FinanceSheet(GoogleSheet):
    def __init__(self, *args, **kwargs):
        super().__init__(*args,**kwargs)
        self._header = self.src[0]
        self._legend = []
        self._stage = {}
        for x in self.src:
            if len(x) > 0:
                self._legend += [ x[0] ]
            else:
                self._legend += [ "" ]


    @property
    def header(self):
        return self._header

    @property
    def legend(self):
        return self._legend

    @property
    def stage(self):
        return self._stage

    @stage.setter
    def stage(self, stage):
        self._stage = stage

    @staticmethod
    def find_ranges(d):
        # d is a dict of <col><row> positions as keys and cell values
        # my sheets tend to have contiguous cells vertically, so that's what we'll try for first
        # anything not in a vertical range can be checked for horizontal range
        # we're not gonna worry about 2d ranges right now
        ranges = []
        # we want a dict where the key is the col and the val is a list of all the rows which
        # have values in that column
        cols = {}
        key_list = d.keys()
        for k in key_list:
            c = ''.join(filter(str.isalpha, k))
            r = ''.join(filter(str.isdigit, k))
            if c not in cols:
                cols[c] = []
            cols[c] += [r]

        for col, rows in cols.items():
            _rows = [ int(x) for x in rows ]
            key_ranges = [list(group) for group in mit.consecutive_groups(_rows)]
            for r in key_ranges:
                first = col + str(r[0])
                vals = [ d[col+str(x)] for x in r ]
                ranges += [ [first, vals] ]
                ranges += [ [first, vals] ]
        return ranges


    def set_stage(self, val, pos=None):
        if pos:
            self.pos = pos
        self.stage[self.pos] = val

    def commit(self):
        print(self.stage)
        results = self.find_ranges(self.stage)
        for pos, r in results:
            self.pos = pos
            self.write(r)
        self.stage = {}


    def col_by_header(self, head):
        self.col = self.all_cols[self.header.index(head)]
        return self.col

    def row_by_legend(self, leg, offset=0):
        self.row = self.legend.index(leg) + 1 + offset
        return self.row

    def set(self, val, col="", row="", offset=0):
        if type(val) == dict:
            new = []
            for x in sorted(val.keys()):
                new += [ val[x] ]
            val = new
        if col:
            self.col_by_header(col)
        if row:
            self.row_by_legend(row, offset=offset)
        self.write(val)
        #self.set_stage(val)

