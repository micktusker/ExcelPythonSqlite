'''
Used to manipulate a Sheet instance from openpyxl
'''
from ExcelReader import ExcelReader
from ExcelColumn import ExcelColumn
class ExcelSheet:
    def __init__(self, sheet):
        '''

        :param sheet:
        '''
        if not sheet.__class__.__name__ == 'Worksheet':
            raise TypeError('Expected type "Worksheet", got %r' %  sheet.__class__.__name__)
        self.sheet = sheet
        self._usedrange_address = self.sheet.dimensions
        self._sheet_values = list(sheet.values)
    @property
    def usedrange_address(self):
        '''
        Return the address of the Used range
        :return: str
        '''
        return self._usedrange_address
    @property
    def column_names(self, column_names_row =  1):
        '''

        :param column_names_row:
        :return: list
        '''
        try:
            return list(self._sheet_values[column_names_row -1])
        except IndexError:
            return []
    @property
    def data_rows(self, first_data_row = 2):
        '''

        :return: list of tuples
        '''
        try:
            return self._sheet_values[first_data_row -1:]
        except IndexError:
            return tuple()
    @property
    def used_row_count(self):
        '''

        :return: int
        '''
        try:
            return len(self._sheet_values)
        except IndexError:
            return 0
    @property
    def used_column_count(self):
        '''

        :return: int
        '''
        try:
            return len(self._sheet_values[0])
        except IndexError:
            return 0
    def get_column_object(self, column_number):
        '''

        :param column_number: int
        :return:
        '''
        try:
            column_name = self.column_names[column_number - 1]
        except IndexError:
            column_name = None
        column_values = []
        if column_number >0 and column_number <= len(self.data_rows[0]):
            for data_row in self.data_rows:
                data_value = self.data_rows[column_number - 1]
                column_values.append(data_value)
        column_object = ExcelColumn(column_name, column_values)
        return column_object
if __name__ == '__main__':
    xl_path = './TestData/Scott.xlsx'
    xlrdr = ExcelReader(xl_path)
    sh = xlrdr.get_sheet_by_name('emp')
    xlsh = ExcelSheet(sh)
    print(xlsh.usedrange_address)
    print(xlsh.column_names)
    print(xlsh.data_rows)
    print(xlsh.get_column_object(1))