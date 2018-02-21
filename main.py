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


class Excelfile(object):

    def __init__(self, file_name):
        self.file_name = file_name
        # Creale clear variables
        self.dictr = {}
        # File params
        self.input_file = None
        self.output_file = None
        self.data_start = None
        # Col params
        self.col_name = None
        self.col_id = None
        self.col_measure = None
        self.col_price = None
        self.col_count = None
        self.col_total = None

        self.read_config()

    def read_xls_file(self):
        file_path = '{0}/{1}'.format(os.getcwd(), self.input_file)
        if os.path.exists(file_path):
            wb = load_workbook(filename=file_path)
            ws = wb.active

            # for col in ws.columns:
            #     pass

            for i in range(self.data_start, ws.max_row - self.data_start):

                if self.col_name != None:
                    name = ws.cell(row=i, column=self.col_name).value
                else:
                    name = ws.cell(row=i, column=2).value

                if self.col_id != None:
                    id = ws.cell(row=i, column=self.col_id).value
                else:
                    id = ws.cell(row=i, column=3).value

                if self.col_price != None:
                    price = ws.cell(row=i, column=self.col_price).value
                else:
                    price = ws.cell(row=i, column=5).value

                if self.col_count != None:
                    count = ws.cell(row=i, column=self.col_count).value
                else:
                    count = ws.cell(row=i, column=6).value

                rec = Record(name, id, price, count)

                rec2 = self.dictr.get(rec.record_id())
                if rec2 != None:
                    rec.count += rec2.count
                    self.dictr.pop(rec2.record_id())

                self.dictr[rec.record_id()] = rec

    def write_xls_file(self):
        file_path = '{0}/{1}'.format(os.getcwd(), self.output_file)
        book = Workbook()
        sheet = book.active

        i = self.data_start
        for item in self.dictr.values():
            if self.col_name != None:
                sheet.cell(row=i, column=self.col_name).value = item.name
            else:
                sheet.cell(row=i, column=2).value = item.name

            if self.col_id != None:
                sheet.cell(row=i, column=self.col_id).value = item.id
            else:
                sheet.cell(row=i, column=3).value = item.id

            if self.col_measure != None:
                sheet.cell(row=i, column=self.col_measure).value = u'шт.'
            else:
                sheet.cell(row=i, column=4).value = u'шт.'

            if self.col_price != None:
                sheet.cell(row=i, column=self.col_price).value = item.price
            else:
                sheet.cell(row=i, column=5).value = item.price

            if self.col_count != None:
                sheet.cell(row=i, column=self.col_count).value = item.price
            else:
                sheet.cell(row=i, column=6).value = item.price

            if self.col_measure != None:
                sheet.cell(row=i, column=self.col_measure).value = item.count * item.price
            else:
                sheet.cell(row=i, column=7).value = item.count * item.price

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

            self.col_name = int(conf['ROWS']['ColName'])
            self.col_id = int(conf['ROWS']['ColId'])
            self.col_measure = int(conf['ROWS']['ColMeasure'])
            self.col_price = int(conf['ROWS']['ColPrice'])
            self.col_count = int(conf['ROWS']['ColCount'])
            self.col_total = int(conf['ROWS']['ColTotal'])


if __name__ == "__main__":
    xls = Excelfile(config_name)
    xls.read_xls_file()
    xls.write_xls_file()
