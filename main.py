import os
# import pandas as pd
# from tabulate import tabulate
from django.contrib.messages.storage import fallback
from openpyxl import Workbook, load_workbook
import configparser
import operator

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
        self.list = []
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
        dict = {}
        bufflist =[]
        file_path = '{0}/{1}'.format(os.getcwd(), self.input_file)
        if os.path.exists(file_path):
            wb = load_workbook(filename=file_path)
            ws = wb.active

            for i in range(self.data_start, ws.max_row - self.data_start):

                rec = Record(
                    ws.cell(row=i, column=self.col_name).value,
                    ws.cell(row=i, column=self.col_id).value,
                    ws.cell(row=i, column=self.col_price).value,
                    ws.cell(row=i, column=self.col_count).value)

                rec2 = dict.get(rec.record_id())
                if rec2 != None:
                    rec.count += rec2.count
                    dict.pop(rec2.record_id())

                dict[rec.record_id()] = rec

            for value in dict.values():
                bufflist.append(value)

            self.list = sorted(bufflist, key=operator.itemgetter(0))



    def write_xls_file(self):
        file_path = '{0}/{1}'.format(os.getcwd(), self.output_file)
        book = Workbook()
        sheet = book.active

        i = self.data_start
        for item in self.list.values():
            sheet.cell(row=i, column=self.col_name).value = item.name
            sheet.cell(row=i, column=self.col_id).value = item.id
            sheet.cell(row=i, column=self.col_measure).value = u'шт.'
            sheet.cell(row=i, column=self.col_price).value = item.price
            sheet.cell(row=i, column=self.col_count).value = item.price
            sheet.cell(row=i, column=self.col_measure).value = item.count * item.price

            i += 1

        book.save(file_path)

    def read_config(self):
        if self.file_name.strip() != '':
            conf = configparser.ConfigParser()
            conf.read(self.file_name.strip())
            self.input_file = conf.get('FILE', 'InputFileName', fallback='input.xls')
            self.output_file = conf('FILE', 'OutputFileName', fallback='output.xls')
            # Start from 0, then title_height + 1
            self.data_start = conf.get('FILE', 'HeaderRowCount', fallback=1) + 1
            #Read row params file
            self.col_name = conf.get('ROWS', 'ColName', fallback=2)
            self.col_id = conf.get('ROWS', 'ColId', fallback=3)
            self.col_measure = conf.get('ROWS', 'ColMeasure', fallback=4)
            self.col_price = conf.get('ROWS', 'ColPrice', fallback=5)
            self.col_count = conf.get('ROWS', 'ColCount', fallback=6)
            self.col_total = conf.get('ROWS', 'ColTotal', fallback=7)


if __name__ == "__main__":
    xls = Excelfile(config_name)
    xls.read_xls_file()
    xls.write_xls_file()
