import q
from bs4 import BeautifulSoup

from ofxstatement.plugin import Plugin
from ofxstatement.parser import StatementParser
from ofxstatement.statement import StatementLine, Statement
from ofxstatement.exceptions import ParseError


class Isracard1Plugin(Plugin):
    """Isracard1 plugin (for developers only)
    """

    def get_parser(self, filename):
        return Isracard1Parser(filename)


class Isracard1Parser(StatementParser):
    def __init__(self, filename):
        self.filename = filename

    def parse(self):
        """Main entry point for parsers

        super() implementation will call to split_records and parse_record to
        process the file.
        """
        with open(self.filename, "r", encoding='iso-8859-8') as f:
            soup = BeautifulSoup(f, 'lxml')
            statement = Statement()
            table = soup.find_all('table', id='trBlueOnWhite12')
            if len(table) == 0:
                raise ParseError(0, "'trBlueonWhite12' table not found")
            q.d()
            return statement
