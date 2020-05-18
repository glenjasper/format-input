#!/usr/bin/env python
# -*- coding: utf-8 -*-
import os
import sys
import time
import argparse
import traceback
import xlsxwriter
import pandas as pd

def menu(args):
    parser = argparse.ArgumentParser(description = "Script que faz a formatação, em arquivos Excel (.xlsx), os arquivos tabulados (.csv) exportadas do Scopus, Web of Science, PubMed, Dimensions ou de um arquivo de texto (.txt)", epilog = "Thank you!")
    parser.add_argument("-t", "--type_file", choices = ofi.ARRAY_TYPE, required = True, type = str.lower, help = ofi.mode_information(ofi.ARRAY_TYPE, ofi.ARRAY_DESCRIPTION))
    parser.add_argument("-i", "--input_file", required = True, help = "Arquivo exportado ou de texto que contem a lista dos DOIs")
    parser.add_argument("-o", "--output", help = "Pasta de saida com a formatação nova")
    parser.add_argument("--version", action = "version", version = "%s %s" % ('%(prog)s', ofi.VERSION))
    args = parser.parse_args()

    ofi.TYPE_FILE = args.type_file
    file_name = os.path.basename(args.input_file)
    file_path = os.path.dirname(args.input_file)
    if file_path is None or file_path == "":
        file_path = os.getcwd().strip()

    ofi.INPUT_FILE = os.path.join(file_path, file_name)
    if not ofi.check_path(ofi.INPUT_FILE):
        ofi.show_print("%s: error: the file '%s' doesn't exist" % (os.path.basename(__file__), ofi.INPUT_FILE), showdate = False)
        ofi.show_print("%s: error: the following arguments are required: -i/--input_file" % os.path.basename(__file__), showdate = False)
        exit()

    if args.output is not None:
        output_name = os.path.basename(args.output)
        output_path = os.path.dirname(args.output)
        if output_path is None or output_path == "":
            output_path = os.getcwd().strip()

        ofi.OUTPUT_PATH = os.path.join(output_path, output_name)
        created = ofi.create_directory(ofi.OUTPUT_PATH)
        if not created:
            ofi.show_print("%s: error: Couldn't create folder '%s'" % (os.path.basename(__file__), ofi.OUTPUT_PATH), showdate = False)
            exit()
    else:
        ofi.OUTPUT_PATH = os.getcwd().strip()
        ofi.OUTPUT_PATH = os.path.join(ofi.OUTPUT_PATH, 'output_format')
        ofi.create_directory(ofi.OUTPUT_PATH)

class FormatInput:

    def __init__(self):
        self.VERSION = 1.0

        self.INPUT_FILE = None
        self.TYPE_FILE = None
        self.OUTPUT_PATH = None

        self.ROOT_DIR = os.path.dirname(os.path.realpath(__file__))
        self.LOG_NAME = "run_%s_%s.log" % (os.path.splitext(os.path.basename(__file__))[0], time.strftime('%Y%m%d'))
        self.LOG_FILE = None

        # Menu
        self.TYPE_SCOPUS = "scopus"
        self.TYPE_WOS = "wos"
        self.TYPE_PUBMED = "pubmed"
        self.TYPE_DIMENSIONS = "dimensions"
        self.TYPE_TXT = "txt"
        self.DESCRIPTION_SCOPUS = "Tipo de arquivo exportado do Scopus (.csv)"
        self.DESCRIPTION_WOS = "Tipo de arquivo exportado do Web of Science (.csv)"
        self.DESCRIPTION_PUBMED = "Tipo de arquivo exportado do PubMed (.csv)"
        self.DESCRIPTION_DIMENSIONS = "Tipo de arquivo exportado do Dimensions (.csv)"
        self.DESCRIPTION_TXT = "Tipo de arquivo .txt"
        self.ARRAY_TYPE = [self.TYPE_SCOPUS, self.TYPE_WOS, self.TYPE_PUBMED, self.TYPE_DIMENSIONS, self.TYPE_TXT]
        self.ARRAY_DESCRIPTION = [self.DESCRIPTION_SCOPUS, self.DESCRIPTION_WOS, self.DESCRIPTION_PUBMED, self.DESCRIPTION_DIMENSIONS, self.DESCRIPTION_TXT]

        # Scopus
        self.scopus_col_authors = 'Authors'
        self.scopus_col_title = 'Title'
        self.scopus_col_year = 'Year'
        self.scopus_col_doi = 'DOI'
        self.scopus_col_document_type = 'Document Type'
        self.scopus_col_languaje = 'Language of Original Document'
        # self.scopus_col_access_type = 'Access Type'
        # self.scopus_col_source = 'Source'

        # Web of Science (WoS)
        self.wos_col_authors = 'AU'
        self.wos_col_title = 'TI'
        self.wos_col_year = 'PY'
        self.wos_col_doi = 'DI'
        self.wos_col_document_type = 'DT'
        self.wos_col_languaje = 'LA'

        # PubMed
        self.pubmed_col_authors = 'Description'
        self.pubmed_col_title = 'Title'
        self.pubmed_col_year = 'Properties'
        self.pubmed_col_doi = 'Details'
        self.pubmed_col_document_type = '' # Doesn't exist
        self.pubmed_col_languaje = '' # Doesn't exist

        # Dimensions
        self.dimensions_col_authors = 'Authors'
        self.dimensions_col_title = 'Title'
        self.dimensions_col_year = 'PubYear'
        self.dimensions_col_doi = 'DOI'
        self.dimensions_col_document_type = 'Publication Type'
        self.dimensions_col_languaje = '' # Doesn't exist

        # Xls Summary
        self.XLS_FILE = 'input_<type>.xlsx'
        self.XLS_SHEET_DETAIL = 'Detail'
        self.XLS_SHEET_WITHOUT_DOI = 'Without DOI'
        self.XLS_SHEET_REDUNDANCIES = 'Redundancies'

        # Xls Columns
        self.xls_col_item = 'Item'
        self.xls_col_title = 'Title'
        self.xls_col_year = 'Year'
        self.xls_col_doi = 'DOI'
        self.xls_col_document_type = 'Document Type'
        self.xls_col_languaje = 'Language'
        self.xls_col_authors = 'Author(s)'

        self.xls_col_redundancy_type = 'Redundancy Type'
        self.xls_val_by_doi = 'By DOI'
        self.xls_val_by_title = 'By Title'

        self.xls_columns_csv = [self.xls_col_item,
                                self.xls_col_title,
                                self.xls_col_year,
                                self.xls_col_doi,
                                self.xls_col_document_type,
                                self.xls_col_languaje,
                                self.xls_col_authors]

        self.xls_columns_txt = [self.xls_col_item,
                                self.xls_col_doi]

        # Fonts
        self.RED = '\033[31m'
        self.GREEN = '\033[32m'
        self.YELLOW = '\033[33m'
        self.BIRED = '\033[1;91m'
        self.BIGREEN = '\033[1;92m'
        self.END = '\033[0m'

    def show_print(self, message, logs = None, showdate = True, font = None):
        msg_print = message
        msg_write = message

        if font is not None:
            msg_print = "%s%s%s" % (font, msg_print, self.END)

        if showdate is True:
            _time = time.strftime('%Y-%m-%d %H:%M:%S')
            msg_print = "%s %s" % (_time, msg_print)
            msg_write = "%s %s" % (_time, message)

        print(msg_print)
        if logs is not None:
            for log in logs:
                if log is not None:
                    with open(log, 'a') as f:
                        f.write("%s\n" % msg_write)
                        f.close()

    def start_time(self):
        return time.time()

    def finish_time(self, start, message = None):
        finish = time.time()
        runtime = time.strftime("%H:%M:%S", time.gmtime(finish - start))
        if message is None:
            return runtime
        else:
            return "%s: %s" % (message, runtime)

    def create_directory(self, path):
        output = True
        try:
            if len(path) > 0 and not os.path.exists(path):
                os.makedirs(path)
        except Exception as e:
            output = False
        return output

    def check_path(self, path):
        _check = False
        if path is not None:
            if len(path) > 0 and os.path.exists(path):
                _check = True
        return _check

    def mode_information(self, array1, array2):
        _information = ["%s: %s" % (i, j) for i, j in zip(array1, array2)]
        return " | ".join(_information)

    def read_txt(self, file):
        content = open(file, 'r').readlines()

        collect_unique = {}
        collect_redundant_doi = {}
        nr_doi = []
        index = 1
        for idx, line in enumerate(content, start = 1):
            line = line.strip()
            if line != '':
                flag_unique = False
                doi = line.lower()
                if doi not in nr_doi:
                    nr_doi.append(doi)
                    flag_unique = True

                collect = {}
                collect[self.xls_col_doi] = doi

                if flag_unique:
                    collect_unique.update({index: collect})
                    index += 1
                else:
                    collect[self.xls_col_redundancy_type] = self.xls_val_by_doi
                    collect_redundant_doi.update({idx: collect})

        collect_papers = {self.XLS_SHEET_DETAIL: collect_unique,
                          self.XLS_SHEET_REDUNDANCIES: collect_redundant_doi}

        return collect_papers

    def read_csv(self, file):
        if self.TYPE_FILE == self.TYPE_SCOPUS:
            separator = ','
            _col_doi = self.scopus_col_doi
        elif self.TYPE_FILE == self.TYPE_WOS:
            separator = '\t'
            _col_doi = self.wos_col_doi
        elif self.TYPE_FILE == self.TYPE_PUBMED:
            separator = ','
            _col_doi = self.pubmed_col_doi
        elif self.TYPE_FILE == self.TYPE_DIMENSIONS:
            separator = ','
            _col_doi = self.dimensions_col_doi

        df = pd.read_csv(filepath_or_buffer = file, sep = separator, header = 0, index_col = False)
        df = df.where(pd.notnull(df), None)

        # Get DOIs
        collect_unique_doi = {}
        collect_redundant_doi = {}
        collect_without_doi = {}
        nr_doi = []
        for idx, row in df.iterrows():
            flag_unique = False
            flag_redundant_doi = False
            flag_null = False

            doi = row[_col_doi]
            if self.TYPE_FILE == self.TYPE_PUBMED:
                info = doi.split('doi:')
                if len(info) > 1:
                    info = info[1].split()
                    doi = info[0]
                else:
                    doi = None

                info = row[self.pubmed_col_year]
                info = info.split('create date:')
                if len(info) > 1:
                    info = info[1].split('/')
                    year = int(info[0])
                else:
                    year = None

            if doi is not None:
                doi = doi.strip()
                doi = doi.lower()
                doi = doi[:-1] if doi.endswith('.') else doi
                if doi not in nr_doi:
                    nr_doi.append(doi)
                    flag_unique = True
                else:
                    flag_redundant_doi = True
            else:
                flag_null = True

            collect = {}
            if self.TYPE_FILE == self.TYPE_SCOPUS:
                collect[self.xls_col_authors] = row[self.scopus_col_authors].strip() if row[self.scopus_col_authors] is not None else row[self.scopus_col_authors]
                collect[self.xls_col_title] = row[self.scopus_col_title].strip() if row[self.scopus_col_title] is not None else row[self.scopus_col_title]
                collect[self.xls_col_year] = row[self.scopus_col_year]
                collect[self.xls_col_doi] = doi
                collect[self.xls_col_document_type] = row[self.scopus_col_document_type].strip() if row[self.scopus_col_document_type] is not None else row[self.scopus_col_document_type]
                collect[self.xls_col_languaje] = row[self.scopus_col_languaje].strip() if row[self.scopus_col_languaje] is not None else row[self.scopus_col_languaje]
            elif self.TYPE_FILE == self.TYPE_WOS:
                collect[self.xls_col_authors] = row[self.wos_col_authors].strip() if row[self.wos_col_authors] is not None else row[self.wos_col_authors]
                collect[self.xls_col_title] = row[self.wos_col_title].strip() if row[self.wos_col_title] is not None else row[self.wos_col_title]
                collect[self.xls_col_year] = row[self.wos_col_year]
                collect[self.xls_col_doi] = doi
                collect[self.xls_col_document_type] = row[self.wos_col_document_type].strip() if row[self.wos_col_document_type] is not None else row[self.wos_col_document_type]
                collect[self.xls_col_languaje] = row[self.wos_col_languaje].strip() if row[self.wos_col_languaje] is not None else row[self.wos_col_languaje]
            elif self.TYPE_FILE == self.TYPE_PUBMED:
                collect[self.xls_col_authors] = row[self.pubmed_col_authors].strip() if row[self.pubmed_col_authors] is not None else row[self.pubmed_col_authors]
                collect[self.xls_col_title] = row[self.pubmed_col_title].strip() if row[self.pubmed_col_title] is not None else row[self.pubmed_col_title]
                collect[self.xls_col_year] = year
                collect[self.xls_col_doi] = doi
                collect[self.xls_col_document_type] = None
                collect[self.xls_col_languaje] = None
            elif self.TYPE_FILE == self.TYPE_DIMENSIONS:
                collect[self.xls_col_authors] = row[self.dimensions_col_authors].strip() if row[self.dimensions_col_authors] is not None else row[self.dimensions_col_authors]
                collect[self.xls_col_title] = row[self.dimensions_col_title].strip() if row[self.dimensions_col_title] is not None else row[self.dimensions_col_title]
                collect[self.xls_col_year] = row[self.dimensions_col_year]
                collect[self.xls_col_doi] = doi
                collect[self.xls_col_document_type] = row[self.dimensions_col_document_type].strip() if row[self.dimensions_col_document_type] is not None else row[self.dimensions_col_document_type]
                collect[self.xls_col_languaje] = None

            if flag_unique:
                collect_unique_doi.update({idx + 1: collect})
            if flag_redundant_doi:
                collect[self.xls_col_redundancy_type] = self.xls_val_by_doi
                collect_redundant_doi.update({idx + 1: collect})
            if flag_null:
                collect_without_doi.update({idx + 1: collect})

        # Get titles
        collect_unique = {}
        collect_redundant_title = {}
        nr_title = []
        index = 1
        for idx, row in collect_unique_doi.items():
            flag_unique = False

            title = row[self.xls_col_title]
            if title is not None:
                title = title.strip()
                title = title.lower()
                title = title[:-1] if title.endswith('.') else title
                if title not in nr_title:
                    nr_title.append(title)
                    flag_unique = True
            else:
                flag_unique = True

            if flag_unique:
                collect_unique.update({index: row})
                index += 1
            else:
                row[self.xls_col_redundancy_type] = self.xls_val_by_title
                collect_redundant_title.update({idx: row})

        collect_redundant = {}
        collect_redundant = collect_redundant_doi.copy()
        collect_redundant.update(collect_redundant_title)
        collect_redundant = {item[0]: item[1] for item in sorted(collect_redundant.items())}

        collect_papers = {self.XLS_SHEET_DETAIL: collect_unique,
                          self.XLS_SHEET_WITHOUT_DOI: collect_without_doi,
                          self.XLS_SHEET_REDUNDANCIES: collect_redundant}

        return collect_papers

    def save_summary_xls(self, data_paper):

        def create_sheet(oworkbook, sheet_type, dictionary, styles_title, styles_rows):
            if self.TYPE_FILE == self.TYPE_TXT:
                _xls_columns = self.xls_columns_txt.copy()
            else:
                _xls_columns = self.xls_columns_csv.copy()

            if sheet_type == self.XLS_SHEET_REDUNDANCIES:
                _xls_columns.append(self.xls_col_redundancy_type)

            _last_col = len(_xls_columns) - 1

            worksheet = oworkbook.add_worksheet(sheet_type)
            worksheet.freeze_panes(row = 1, col = 0) # Freeze the first row.
            worksheet.autofilter(first_row = 0, first_col = 0, last_row = 0, last_col = _last_col) # 'A1:H1'
            worksheet.set_default_row(height = 14.5)

            # Add columns
            for icol, column in enumerate(_xls_columns):
                worksheet.write(0, icol, column, styles_title)

            # Add rows
            if self.TYPE_FILE == self.TYPE_TXT:
                worksheet.set_column(first_col = 0, last_col = 0, width = 7)  # Column A:A
                worksheet.set_column(first_col = 1, last_col = 1, width = 33) # Column B:B
                if sheet_type == self.XLS_SHEET_REDUNDANCIES:
                    worksheet.set_column(first_col = 2, last_col = 2, width = 19) # Column C:C
            else:
                worksheet.set_column(first_col = 0, last_col = 0, width = 7)  # Column A:A
                worksheet.set_column(first_col = 1, last_col = 1, width = 40) # Column B:B
                worksheet.set_column(first_col = 2, last_col = 2, width = 8)  # Column C:C
                worksheet.set_column(first_col = 3, last_col = 3, width = 33) # Column D:D
                worksheet.set_column(first_col = 4, last_col = 4, width = 18) # Column E:E
                worksheet.set_column(first_col = 5, last_col = 5, width = 12) # Column F:F

                worksheet.set_column(first_col = 6, last_col = 6, width = 36) # Column G:G
                if sheet_type == self.XLS_SHEET_REDUNDANCIES:
                    worksheet.set_column(first_col = 7, last_col = 7, width = 19) # Column H:H

            icol = 0
            for irow, (index, item) in enumerate(dictionary.items(), start = 1):
                col_doi = item[self.xls_col_doi]
                if sheet_type == self.XLS_SHEET_REDUNDANCIES:
                    redundancy_type = item[self.xls_col_redundancy_type]

                if self.TYPE_FILE == self.TYPE_TXT:
                    worksheet.write(irow, icol + 0, index, styles_rows)
                    worksheet.write(irow, icol + 1, col_doi, styles_rows)
                    if sheet_type == self.XLS_SHEET_REDUNDANCIES:
                        worksheet.write(irow, icol + 2, redundancy_type, styles_rows)
                else:
                    worksheet.write(irow, icol + 0, index, styles_rows)
                    worksheet.write(irow, icol + 1, item[self.xls_col_title], styles_rows)
                    worksheet.write(irow, icol + 2, item[self.xls_col_year], styles_rows)
                    worksheet.write(irow, icol + 3, col_doi, styles_rows)
                    worksheet.write(irow, icol + 4, item[self.xls_col_document_type], styles_rows)
                    worksheet.write(irow, icol + 5, item[self.xls_col_languaje], styles_rows)
                    worksheet.write(irow, icol + 6, item[self.xls_col_authors], styles_rows)
                    if sheet_type == self.XLS_SHEET_REDUNDANCIES:
                        worksheet.write(irow, icol + 7, redundancy_type, styles_rows)

        workbook = xlsxwriter.Workbook(self.XLS_FILE)

        # Styles
        cell_format_title = workbook.add_format({'bold': True,
                                                 'font_color': 'white',
                                                 'bg_color': 'black',
                                                 'align': 'center',
                                                 'valign': 'vcenter'})
        cell_format_row = workbook.add_format({'text_wrap': True, 'valign': 'top'})

        create_sheet(workbook, self.XLS_SHEET_DETAIL, data_paper[self.XLS_SHEET_DETAIL], cell_format_title, cell_format_row)
        if self.TYPE_FILE != self.TYPE_TXT:
            create_sheet(workbook, self.XLS_SHEET_WITHOUT_DOI, data_paper[self.XLS_SHEET_WITHOUT_DOI], cell_format_title, cell_format_row)
        create_sheet(workbook, self.XLS_SHEET_REDUNDANCIES, data_paper[self.XLS_SHEET_REDUNDANCIES], cell_format_title, cell_format_row)

        workbook.close()

def main(args):
    try:
        start = ofi.start_time()
        menu(args)

        ofi.LOG_FILE = os.path.join(ofi.OUTPUT_PATH, ofi.LOG_NAME)
        ofi.XLS_FILE = os.path.join(ofi.OUTPUT_PATH, ofi.XLS_FILE.replace('<type>', ofi.TYPE_FILE))
        ofi.show_print("#############################################################################", [ofi.LOG_FILE], font = ofi.BIGREEN)
        ofi.show_print("############################### Format Input ################################", [ofi.LOG_FILE], font = ofi.BIGREEN)
        ofi.show_print("#############################################################################", [ofi.LOG_FILE], font = ofi.BIGREEN)

        # Read input file
        input_information = {}
        if ofi.TYPE_FILE == ofi.TYPE_TXT:
            ofi.show_print("Reading the .txt file", [ofi.LOG_FILE], font = ofi.GREEN)
            input_information = ofi.read_txt(ofi.INPUT_FILE)
        elif ofi.TYPE_FILE == ofi.TYPE_SCOPUS:
            ofi.show_print("Reading the .csv file from Scopus", [ofi.LOG_FILE], font = ofi.GREEN)
            input_information = ofi.read_csv(ofi.INPUT_FILE)
        elif ofi.TYPE_FILE == ofi.TYPE_WOS:
            ofi.show_print("Reading the .csv file from Web of Science", [ofi.LOG_FILE], font = ofi.GREEN)
            input_information = ofi.read_csv(ofi.INPUT_FILE)
        elif ofi.TYPE_FILE == ofi.TYPE_PUBMED:
            ofi.show_print("Reading the .csv file from PubMed", [ofi.LOG_FILE], font = ofi.GREEN)
            input_information = ofi.read_csv(ofi.INPUT_FILE)
        elif ofi.TYPE_FILE == ofi.TYPE_DIMENSIONS:
            ofi.show_print("Reading the .csv file from Dimensions", [ofi.LOG_FILE], font = ofi.GREEN)
            input_information = ofi.read_csv(ofi.INPUT_FILE)
        ofi.show_print("Input file: %s" % ofi.INPUT_FILE, [ofi.LOG_FILE])
        ofi.show_print("", [ofi.LOG_FILE])

        ofi.save_summary_xls(input_information)
        ofi.show_print("Output file: %s" % ofi.XLS_FILE, [ofi.LOG_FILE], font = ofi.GREEN)

        ofi.show_print("", [ofi.LOG_FILE])
        ofi.show_print(ofi.finish_time(start, "Elapsed time"), [ofi.LOG_FILE])
        ofi.show_print("Done!", [ofi.LOG_FILE])
    except Exception as e:
        ofi.show_print("\n%s" % traceback.format_exc(), [ofi.LOG_FILE], font = ofi.RED)
        ofi.show_print(ofi.finish_time(start, "Elapsed time"), [ofi.LOG_FILE])
        ofi.show_print("Done!", [ofi.LOG_FILE])

if __name__ == '__main__':
    ofi = FormatInput()
    main(sys.argv)
