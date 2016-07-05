#!/usr/bin/env python3
import q
from bs4 import BeautifulSoup
import dateutil.parser
import openpyxl
import zipfile

from ofxstatement.plugin import Plugin
from ofxstatement.parser import StatementParser
from ofxstatement.statement import StatementLine, Statement
from ofxstatement.exceptions import ParseError


class Hapoalim1Plugin(Plugin):
    """Hapoalim1 plugin (for developers only)
    """

    def get_parser(self, filename):
        return Hapoalim1Parser(filename)

class Hapoalim1Parser(StatementParser):
    def __init__(self, filename):
        self.filename = filename

    def parse_tr(self, tr):
        def get_header(tr, n):
            try:
                return str(tr.find_all(headers='header'+str(n))[0].string)
            except Exception as e:
                q.d()

        def get_float(s):
            NBSP = u'\xa0'
            if s == NBSP:
                return None
            return float(s.replace(',', ''))

        hdrs = tr.find_all(headers=True)
        assert len(hdrs) in [0, 7]
        if len(hdrs) == 0:
            cont = True
            date_or_comm = str(tr.find_all(colspan='5')[0].string)
            desc = None
            amt = None
            bal = None
        else:
            cont = False
            date_or_comm = dateutil.parser.parse(get_header(tr, 1), dayfirst=True)
            desc = get_header(tr, 2)

            paym = get_float(get_header(tr, 5))
            depo = get_float(get_header(tr, 6))

            if (depo is None) == (paym is None):
                raise Exception('Either depo or paym need to exist, but not both: ' + str(tr))

            amt = depo if depo is not None else -paym

            bal = get_float(get_header(tr, 7))

        return cont, date_or_comm, desc, amt, bal

    def detect_version(self):
        def validator_old(filename):
            try:
                with open(filename, 'r', encoding='iso-8859-8') as f:
                    bs = BeautifulSoup(f, 'lxml')
                table = bs.find_all('table', id='mytable_body')[0]
                return table.find_all('th', id='header1')[0].string == "תאריך"
            except (UnicodeDecodeError, IndexError):
                return False
            return False

        def validator_new(filename):
            try:
                with open(filename, 'r', encoding='iso-8859-8') as f:
                    bs = BeautifulSoup(f, 'lxml')
                table = bs.find_all('table', i__d='trBlueOnWhite12')[0]
                return table.tr.td.string == "תאריך"
            except (UnicodeDecodeError, IndexError):
                return False
            return False

        def validator_xslx(filename):
            try:
                with open(filename, 'rb') as f:
                    wb = openpyxl.load_workbook(f)
                if wb.sheetnames[0] != 'גיליון1':
                    return False
                return wb.worksheets[0]['A6'].value == "תאריך"

            except (openpyxl.utils.exceptions.InvalidFileException, zipfile.BadZipFile):
                return False
            return False

        validators = [validator_old, validator_new, validator_xslx]
        for n, validator in enumerate(validators):
            if validator(self.filename):
                return n

    def parse(self):
        v = self.detect_version()
        if v is None:
            raise Exception("Unsupported file %s" % self.filename)
        print("Detected file of version %d" % v)
        return Statement()
#        with open(self.filename, "r", encoding='iso-8859-8') as f:
#            soup = BeautifulSoup(f, 'lxml')
#            statement = Statement(currency='ILS')
#            data = []
#            q.d()
#            for tr in soup.find_all('table', id='mytable_body')[0].find_all(id='TR_ROW_BANKTABLE'):
#                cont, date_or_comm, desc, amt, bal = self.parse_tr(tr)
#                if cont:
#                    data[-1]['comm'] = date_or_comm
#                else:
#                    new_line = {}
#
#                    new_line['date'] = date_or_comm
#                    new_line['description'] = desc
#                    new_line['amount'] = amt
#                    new_line['balance'] = bal
#                    new_line['comm'] = None
#
#                    data.append(new_line)
#
#            for d in data:
#                stmt_line = StatementLine(date=d['date'], amount=d['amount'])
#                #stmt_line.end_balance = d['balance']
#                stmt_line.payee = d['description']
#                if d['comm'] is not None:
#                    stmt_line.memo = d['comm']
#                stmt_line.trntype = "CASH" if d['amount'] < 0 else "DEP"
#                print(stmt_line)
#                #q.d()
#                statement.lines.append(stmt_line)
#            return statement
