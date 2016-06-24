from ofxstatement.plugin import Plugin
from ofxstatement.parser import StatementParser
from ofxstatement.statement import StatementLine


class Isracard1Plugin(Plugin):
    """Isracard1 plugin (for developers only)
    """

    def getParser(self, filename):
        return Isracard1Parser(filename)


class Isracard1Parser(StatementParser):
    def __init__(self, filename):
        self.filename = filename

    def parse(self):
        """Main entry point for parsers

        super() implementation will call to split_records and parse_record to
        process the file.
        """
        with open(self.filename, "r") as f:
            self.input = f
            return super(Isracard1Parser, self).parse()

    def split_records(self):
        """Return iterable object consisting of a line per transaction
        """
        return []

    def parse_record(self, line):
        """Parse given transaction line and return StatementLine object
        """
        return StatementLine()
