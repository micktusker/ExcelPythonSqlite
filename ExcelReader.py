import openpyxl
import os.path
class ExcelReader:
    def __init__(self, path_to_excel_file):
        '''
        Initialize with a path to an Excel file. Check that file exists or throw an exception.
        :param path_to_excel_file:
        '''
        if not os.path.isfile(path_to_excel_file):
            raise IOError('Excel file %r not found' % path_to_excel_file)
        self.path_to_excel_file = path_to_excel_file
        self.wb = openpyxl.load_workbook(self.path_to_excel_file)
        self._sheet_names = self.wb.get_sheet_names()
        self._sheet_count = len(self._sheet_names)
    @property
    def sheet_count(self):
        '''Return the number of sheets in the given Excel file as an integer.'''
        return self._sheet_count
    @property
    def sheet_names(self):
        '''Return a list with the sheet names'''
        return self._sheet_names
    def get_sheet_by_name(self, sheet_name):
        '''
        Given a sheet name return a Worksheet object.
        Return None if the given sheet name does not exist.
        :param sheet_name:
        :return: Worksheet
        '''
        try:
            sheet = self.wb.get_sheet_by_name(sheet_name)
            return sheet
        except KeyError:
            return None

if __name__ == '__main__':
    xl_path = './TestData/Scott.xlsx'
    xlrdr = ExcelReader(xl_path)
    print('\n'.join(xlrdr.sheet_names))
    print(xlrdr.get_sheet_by_name('emp').__class__.__name__)
    sh = xlrdr.get_sheet_by_name('emp')
    #print('\n'.join(dir(sh)))
    print(sh.dimensions)
    first_row_cells = sh.rows.__next__()
    print('\n'.join(dir(first_row_cells[0])))
    for value in sh.values:
        print(value)