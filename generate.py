import openpyxl
import styler
from openpyxl.utils import coordinate_from_string, column_index_from_string, get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from datetime import datetime


def a1_to_rc(a1):
    xy = coordinate_from_string(a1)
    col = column_index_from_string(xy[0])
    row = xy[1]
    return row, col


class WorkBookMaker:
    def __init__(self, fname_markdown, fname_xls):
        with open(fname_markdown) as f:
            self.lines = f.readlines()
        self.fname_xls = fname_xls
        self.wb = openpyxl.Workbook()
        for sheetName in self.wb.sheetnames:
            del self.wb[sheetName]
        self.scalars = {}
        self.ranges = {}
        self.formulas = []
        self.r = 1
        self.c = 1
        self.ws = None
        self.line_number = -1

    def process(self):
        while self.line_number < (len(self.lines) - 1):
            self.line_number += 1
            line = self.lines[self.line_number]

            # a new sheet
            if line.startswith("# "):
                sheet_name = line[2:]
                self.r = 1
                self.c = 1
                self.wb.create_sheet(sheet_name)
                self.ws = self.wb[sheet_name]
            elif line.startswith("## "):
                value = line[3:]
                self.__cell().value = value
                styler.heading(self.__cell())
                self.r += 1
            elif self.__try_use_cell(line):
                pass
            # a block of markdown text
            elif line.startswith("[text]"):
                self.__process_text()
            elif len(line.strip()) == 0:
                self.r += 1
            # a table of columns, will continue to [endtable_c]
            elif line.startswith("[table_c]"):
                self.__process_table_c()
            else:
                self.__add_values(line)
        self.__update_formulas()

    def save(self):
        self.wb.save(fname_xls)

    def __update_formulas(self):
        """ replace variable names with references
        """
        for f in self.formulas:
            value = f.formula
            for key in self.scalars.keys():
                if value.find(key) > -1:
                    value = value.replace(key, self.scalars[key])
                    print(f"replacing {key} with {self.scalars[key]} : {value} ")
            ws = self.wb[f.sheet]
            ws.cell(f.row, f.col, value)

    def __add_values(self, line):
        first = line.find(',')
        if first < 0:
            self.__line_error(line)
        part0 = line[:first]
        part1 = line[first + 1:]
        cell = self.__cell()
        cell.value = part0
        styler.label(cell)
        cell = self.__cell(1)
        if part1.strip()[0] == "=":
            print(f"formula: {part1}")
            self.formulas.append(Formula(self.ws.title, self.r, self.c + 1, part1))
            styler.formula(cell)
        else:
            self.__set_value(cell, part1)
            self.scalars[part0] = f"{get_column_letter(self.c + 1)}{self.r}"
            styler.value(cell)
        self.r += 1

    def __set_value(self, cell, str):
        str = str.strip()
        if str.endswith('%'):
            value = float(str[:-1]) / 100
            cell.value = value
            cell.number_format = '0.00%'
            return
        parts = str.split('-')
        if len(parts) == 3:
            value = v = datetime(int(parts[0]), int(parts[1]), int(parts[2]))
            cell.value = value
            cell.number_format = "YYYY-MM-DD"
            return
        if str.startswith('['):
            i0 = str.find('[')
            i1 = str.find(']')
            list_str = str[i0:i1]
            list_vals = list_str.split(',')
            value = str[i1 + 1:]
            dv = DataValidation(type="list", formula1=f'"{list_str}"', allow_blank=False)
            dv.add(cell)
            self.ws.add_data_validation(dv)
            cell.value = value
            return
        cell.value = float(str)
        cell.number_format = "0.00"

    def __line_error(self, line):
        print(f"Can not process: {line}")

    def __process_text(self):
        self.line_number += 1
        line = self.lines[self.line_number]
        while not line.startswith("[endtext]"):
            cell = self.ws.cell(row=self.r, column=self.c)
            cell.value = line
            styler.text(cell)
            self.line_number += 1
            self.r += 1
            line = self.lines[self.line_number]
        self.ws.column_dimensions[get_column_letter(self.c)].width = 100

    def __cell(self, col_offset=None):
        if col_offset is None:
            c = 0
        else:
            c = col_offset
        return self.ws.cell(row=self.r, column=self.c + c)

    def __try_use_cell(self, line):
        if not line.startswith('['):
            return False
        test = line.replace('[', '')
        test = test.replace(']', '')
        test = test.replace(' ', '')
        try:
            xy = coordinate_from_string(test)
            self.r = xy[1]
            self.c = column_index_from_string(xy[0])
            return True
        except:
            return False

    def __process_table_c(self):
        pass


class Formula:
    def __init__(self, sheet, row, col, formula):
        self.sheet = sheet
        self.row = row
        self.col = col
        self.formula = formula


fname_markdown = "C:/Dev/python/xls_markdown/caplet.md"
fname_xls = "C:/Dev/python/xls_markdown/caplet_test.xlsx"
wbm = WorkBookMaker(fname_markdown, fname_xls)
wbm.process()
wbm.save()
