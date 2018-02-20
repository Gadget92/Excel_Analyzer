import os
# import pandas as pd
# from tabulate import tabulate
from openpyxl import Workbook, load_workbook
import configparser

config_name = "config.ini"

class Record(object):

    def __init__(self, name, id, price, count):
        self.name = name
        self.price = price
        self.count = count
        self.id = id

    def record_id(self):
        return '{0}_{1}'.format(self.id, self.price)

class Exclfile(object):

    def __init__(self, file_name):
        self.file_name = file_name
        # Creale clear variables
        self.dictr = {}
        self.input_file = ''
        self.output_file = ''
        self.data_start = 0


    def read_xls_file(self):
        file_path = '{0}/{1}'.format(os.getcwd(), self.input_file)
        if os.path.exists(file_path):
            wb = load_workbook(filename=file_path)
            ws = wb.active

            # for col in ws.columns:
            #     pass

            for i in range(self.data_start, ws.max_row - self.data_start):
                name = ws.cell(row=i, column=2).value
                id = ws.cell(row=i, column=3).value
                price = ws.cell(row=i, column=5).value
                count = ws.cell(row=i, column=6).value
                rec = Record(name, id, price, count)

                rec2 = self.dictr.get(rec.record_id())
                if rec2 != None :
                    rec.count += rec2.count
                    self.dictr.pop(rec2.record_id())

                self.dictr[rec.record_id()] = rec


    def write_xls_file(self):
        file_path = '{0}/{1}'.format(os.getcwd(), self.output_file)
        book = Workbook()
        sheet = book.active

        i = self.data_start
        for item in self.dictr.values():
            sheet.cell(row=i, column=2).value = item.name
            sheet.cell(row=i, column=3).value = item.id
            sheet.cell(row=i, column=4).value = u'шт.'
            sheet.cell(row=i, column=5).value = item.price
            sheet.cell(row=i, column=6).value = item.count
            sheet.cell(row=i, column=7).value = '''=MMULT(E{0};F{0})'''.format(i)
            i += 1

        book.save(file_path)

    def read_config(self):
        if self.file_name.strip() != '':
            conf = configparser.ConfigParser()
            conf.read(self.file_name.strip())
            self.input_file = conf['FILE']['InputFileName']
            self.output_file = conf['FILE']['OutputFileName']
            # Start from 0, then title_height + 1
            self.data_start = int(conf['FILE']['HeaderRowCount']) + 1



if __name__ == "__main__":
    xls = Exclfile(config_name)

    xls.read_config()
    xls.read_xls_file()
    xls.write_xls_file()
