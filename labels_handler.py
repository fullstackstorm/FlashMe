"""
Labels Handler: coded by @jjonamos
Read/writes label definitions from/to a given Excel file,
which must contain a sheet named 'Py_Variables', which in
turn must contain a table called 'Label_Dictionary' with
headers 'Key' and 'Value'. Creates a dictionary list for
easy parsing elsewhere.
"""
import xlwings

class label:
    def __init__(self, excel_file):
        self._work_book = xlwings.Book(excel_file)
        work_sheet = self._work_book.sheets("Py_Variables")
        self._key_list = work_sheet["Label_Dictionary[Key]"]
        self._value_list = work_sheet["Label_Dictionary[Value]"]
        self.dictionary = {}
        self.__init_dictionary()

    def __init_dictionary(self):
        with xlwings.App(visible = False):
            key_list = [str(entry.value) for entry in self._key_list]
            value_list = [str(entry.value) for entry in self._value_list]
            for key, value in zip(key_list, value_list): self.dictionary[key] = value

    def save(self):
        with xlwings.App(visible = False):
            self._key_list.options(transpose = True).value = list(self.dictionary.keys())
            self._value_list.options(transpose = True).value = list(self.dictionary.values())
            self._work_book.save()