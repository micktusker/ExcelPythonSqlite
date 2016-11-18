
class ExcelColumn:
    def __init__(self, name, values):
        '''
        Initialize with a string and a list of values
        :param name: str
        :param values: list
        '''
        self._name = name
        self._values = values
    def name(self):
        return self._name
    def values(self):
        return self._values
    def get_column_type(self):
       pass