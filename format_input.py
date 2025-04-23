#!/usr/bin/env python
# -*- coding: utf-8 -*-
import os
import re
import sys
import time
import shutil
import argparse
import tempfile
import traceback
import xlsxwriter
import numpy as np
import pandas as pd
from colorama import init
from pprint import pprint
init()

def menu():
    parser = argparse.ArgumentParser(description = "This script reads the exported (.csv|.txt) files from Scopus, Web of Science, PubMed, PubMed Central, Dimensions, Cochrane, Embase, ScienceDirect, IEEE, BVS, CAB, SciELO, or Google Scholar (exported from Publish or Perish) databases and turns each of them into a new file with an unique format. This script will ignore duplicated records.", epilog = "Thank you!")
    parser.add_argument("-t", "--type_file", choices = ofi.ARRAY_TYPE, required = True, type = str.lower, help = ofi.mode_information(ofi.ARRAY_TYPE, ofi.ARRAY_DESCRIPTION))
    parser.add_argument("-i", "--input_file", required = True, help = "Input file .csv or .txt")
    parser.add_argument("-o", "--output", help = "Output folder")
    parser.add_argument("--version", action = "version", version = "%s %s" % ('%(prog)s', ofi.VERSION))
    args = parser.parse_args()

    ofi.TYPE_FILE = args.type_file
    file_name = os.path.basename(args.input_file)
    file_path = os.path.dirname(args.input_file)
    if file_path is None or file_path == "":
        file_path = os.getcwd().strip()

    ofi.INPUT_FILE = os.path.join(file_path, file_name)
    if not ofi.check_path(ofi.INPUT_FILE):
        ofi.show_print("%s: error: the file '%s' doesn't exist" % (os.path.basename(__file__), ofi.INPUT_FILE), showdate = False, font = ofi.YELLOW)
        ofi.show_print("%s: error: the following arguments are required: -i/--input_file" % os.path.basename(__file__), showdate = False, font = ofi.YELLOW)
        exit()

    if args.output:
        output_name = os.path.basename(args.output)
        output_path = os.path.dirname(args.output)
        if output_path is None or output_path == "":
            output_path = os.getcwd().strip()

        ofi.OUTPUT_PATH = os.path.join(output_path, output_name)
        created = ofi.create_directory(ofi.OUTPUT_PATH)
        if not created:
            ofi.show_print("%s: error: Couldn't create folder '%s'" % (os.path.basename(__file__), ofi.OUTPUT_PATH), showdate = False, font = ofi.YELLOW)
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
        self.TYPE_PUBMED_CENTRAL = "pmc"
        self.TYPE_DIMENSIONS = "dimensions"
        self.TYPE_GOOGLE_SCHOLAR = "scholar"
        self.TYPE_COCHRANE = "cochrane"
        self.TYPE_EMBASE = "embase"
        self.TYPE_SCIENCEDIRECT = "sciencedirect"
        self.TYPE_IEEE = "ieee"
        self.TYPE_BVS = "bvs"
        self.TYPE_CAB = "cab"
        self.TYPE_SCIELO = "scielo"
        self.TYPE_TXT = "txt"
        self.DESCRIPTION_SCOPUS = "Indicates that the file (.csv) was exported from Scopus"
        self.DESCRIPTION_WOS = "Indicates that the file (.csv) was exported from Web of Science"
        self.DESCRIPTION_PUBMED = "Indicates that the file (.csv) was exported from PubMed"
        self.DESCRIPTION_PUBMED_CENTRAL = "Indicates that the file (.txt) was exported from PubMed Central, necessarily in MEDLINE format"
        self.DESCRIPTION_DIMENSIONS = "Indicates that the file (.csv) was exported from Dimensions"
        self.DESCRIPTION_GOOGLE_SCHOLAR = "Indicates that the file (.csv) was exported from Publish or Perish (Google Scholar option)"
        self.DESCRIPTION_COCHRANE = "Indicates that the file (.csv) was exported from Cochrane"
        self.DESCRIPTION_EMBASE = "Indicates that the file (.csv) was exported from Embase"
        self.DESCRIPTION_SCIENCEDIRECT = "Indicates that the file (.ris) was exported from ScienceDirect"
        self.DESCRIPTION_IEEE = "Indicates that the file (.csv) was exported from IEEE"
        self.DESCRIPTION_BVS = "Indicates that the file (.csv) was exported from BVS"
        self.DESCRIPTION_CAB = "Indicates that the file (.csv) was exported from CAB"
        self.DESCRIPTION_SCIELO = "Indicates that the file (.csv) was exported from SciELO"
        self.DESCRIPTION_TXT = "Indicates that it is a text file (.txt)"
        self.ARRAY_TYPE = [self.TYPE_SCOPUS,
                           self.TYPE_WOS,
                           self.TYPE_PUBMED,
                           self.TYPE_PUBMED_CENTRAL,
                           self.TYPE_DIMENSIONS,
                           self.TYPE_GOOGLE_SCHOLAR,
                           self.TYPE_COCHRANE,
                           self.TYPE_EMBASE,
                           self.TYPE_SCIENCEDIRECT,
                           self.TYPE_IEEE,
                           self.TYPE_BVS,
                           self.TYPE_CAB,
                           self.TYPE_SCIELO,
                           self.TYPE_TXT]
        self.ARRAY_DESCRIPTION = [self.DESCRIPTION_SCOPUS,
                                  self.DESCRIPTION_WOS,
                                  self.DESCRIPTION_PUBMED,
                                  self.DESCRIPTION_PUBMED_CENTRAL,
                                  self.DESCRIPTION_DIMENSIONS,
                                  self.DESCRIPTION_GOOGLE_SCHOLAR,
                                  self.DESCRIPTION_COCHRANE,
                                  self.DESCRIPTION_EMBASE,
                                  self.DESCRIPTION_SCIENCEDIRECT,
                                  self.DESCRIPTION_IEEE,
                                  self.DESCRIPTION_BVS,
                                  self.DESCRIPTION_CAB,
                                  self.DESCRIPTION_SCIELO,
                                  self.DESCRIPTION_TXT]

        # Scopus
        self.scopus_col_authors = 'Authors'
        self.scopus_col_title = 'Title'
        self.scopus_col_year = 'Year'
        self.scopus_col_doi = 'DOI'
        self.scopus_col_document_type = 'Document Type'
        self.scopus_col_language = 'Language of Original Document'
        self.scopus_col_cited_by = 'Cited by'
        self.scopus_col_abstract = 'Abstract'
        # self.scopus_col_access_type = 'Access Type'
        # self.scopus_col_source = 'Source'

        # Web of Science (WoS) | SciELO
        self.wos_col_authors = 'AU'
        self.wos_col_title = 'TI'
        self.wos_col_year = 'PY'
        self.wos_col_doi = 'DI'
        self.wos_col_document_type = 'DT'
        self.wos_col_language = 'LA'
        self.wos_col_cited_by = 'TC'
        self.wos_col_abstract = 'AB'

        # PubMed
        self.pubmed_col_authors = 'Authors'
        self.pubmed_col_title = 'Title'
        self.pubmed_col_year = 'Publication Year'
        self.pubmed_col_doi = 'DOI'
        self.pubmed_col_document_type = '' # Doesn't exist
        self.pubmed_col_language = '' # Doesn't exist
        self.pubmed_col_cited_by = '' # Doesn't exist
        self.pubmed_col_abstract = '' # Doesn't exist

        # PubMed Central
        self.pmc_col_authors = 'Authors'
        self.pmc_col_title = 'Title'
        self.pmc_col_year = 'Publication Year'
        self.pmc_col_doi = 'DOI'
        self.pmc_col_document_type = 'Document Type'
        self.pmc_col_language = 'Language'
        self.pmc_col_cited_by = '' # Doesn't exist
        self.pmc_col_abstract = 'Abstract'

        # Dimensions
        self.dimensions_col_authors = 'Authors'
        self.dimensions_col_title = 'Title'
        self.dimensions_col_year = 'PubYear'
        self.dimensions_col_doi = 'DOI'
        self.dimensions_col_document_type = '' # Doesn't exist
        self.dimensions_col_language = '' # Doesn't exist
        self.dimensions_col_cited_by = 'Times cited'
        self.dimensions_col_abstract = 'Abstract'

        # Google Scholar (Publish or Perish)
        self.scholar_col_authors = 'Authors'
        self.scholar_col_title = 'Title'
        self.scholar_col_year = 'Year'
        self.scholar_col_doi = 'DOI'
        self.scholar_col_document_type = '' # Doesn't exist
        self.scholar_col_language = '' # Doesn't exist
        self.scholar_col_cited_by = 'Cites'
        self.scholar_col_abstract = '' # Doesn't exist

        # Cochrane
        self.cochrane_col_authors = 'Author(s)'
        self.cochrane_col_title = 'Title'
        self.cochrane_col_year = 'Year'
        self.cochrane_col_doi = 'DOI'
        self.cochrane_col_document_type = '' # Doesn't exist
        self.cochrane_col_language = '' # Doesn't exist
        self.cochrane_col_cited_by = '' # Doesn't exist
        self.cochrane_col_abstract = 'Abstract'

        # Embase
        self.embase_col_authors = 'Author Names'
        self.embase_col_title = 'Title'
        self.embase_col_year = 'Publication Year'
        self.embase_col_doi = 'DOI'
        self.embase_col_document_type = 'Publication Type'
        self.embase_col_language = 'Article Language'
        self.embase_col_cited_by = '' # Doesn't exist
        self.embase_col_abstract = 'Abstract'

        # ScienceDirect
        self.sciencedirect_col_authors = 'AU'
        self.sciencedirect_col_title = 'T1'
        self.sciencedirect_col_year = 'PY'
        self.sciencedirect_col_doi = 'DO'
        self.sciencedirect_col_document_type = '' # Doesn't exist
        self.sciencedirect_col_language = '' # Doesn't exist
        self.sciencedirect_col_cited_by = '' # Doesn't exist
        self.sciencedirect_col_abstract = 'AB'

        # IEEE
        self.ieee_col_authors = 'Authors'
        self.ieee_col_title = 'Document Title'
        self.ieee_col_year = 'Publication Year'
        self.ieee_col_doi = 'DOI'
        self.ieee_col_document_type = '' # Doesn't exist
        self.ieee_col_language = '' # Doesn't exist
        self.ieee_col_cited_by = '' # Doesn't exist
        self.ieee_col_abstract = 'Abstract'

        # BVS
        self.bvs_col_authors = 'Authors'
        self.bvs_col_title = 'Title'
        self.bvs_col_year = 'Publication year'
        self.bvs_col_doi = 'DOI'
        self.bvs_col_document_type = 'Type'
        self.bvs_col_language = 'Language'
        self.bvs_col_cited_by = '' # Doesn't exist
        self.bvs_col_abstract = 'Abstract'

        # CAB
        self.cab_col_authors = 'Authors'
        self.cab_col_title = 'Title'
        self.cab_col_year = 'Year of Publication'
        self.cab_col_doi = 'Doi'
        self.cab_col_document_type = '' # Doesn't exist
        self.cab_col_language = 'Languages of Text'
        self.cab_col_cited_by = '' # Doesn't exist
        self.cab_col_abstract = 'Abstract Text'

        # Xls Summary
        self.XLS_FILE = 'input_<type>.xlsx'
        self.XLS_SHEET_UNIQUE = 'Unique'
        self.XLS_SHEET_WITHOUT_DOI = 'Without DOI'
        self.XLS_SHEET_DUPLICATES = 'Duplicates'

        # Xls Columns
        self.xls_col_item = 'Item'
        self.xls_col_title = 'Title'
        self.xls_col_abstract = 'Abstract'
        self.xls_col_year = 'Year'
        self.xls_col_doi = 'DOI'
        self.xls_col_document_type = 'Document Type'
        self.xls_col_language = 'Language'
        self.xls_col_cited_by = 'Cited By'
        self.xls_col_authors = 'Author(s)'

        self.xls_col_duplicate_type = 'Duplicate Type'
        self.xls_val_by_doi = 'By DOI'
        self.xls_val_by_title = 'By Title'

        self.xls_columns_csv = [self.xls_col_item,
                                self.xls_col_title,
                                self.xls_col_abstract,
                                self.xls_col_year,
                                self.xls_col_doi,
                                self.xls_col_document_type,
                                self.xls_col_language,
                                self.xls_col_cited_by,
                                self.xls_col_authors]

        self.xls_columns_txt = [self.xls_col_item,
                                self.xls_col_doi]

        # PubMed Central | MEDLINE
        self.MEDLINE_START = ['AB  -',
                              'AD  -',
                              'AID -',
                              'AU  -',
                              'AUID-',
                              'CN  -',
                              'DEP -',
                              'DP  -',
                              'FAU -',
                              'FIR -',
                              'GR  -',
                              'IP  -',
                              'IR  -',
                              'IS  -',
                              'JT  -',
                              'LA  -',
                              'LID -',
                              'MID -',
                              'OAB -',
                              'OABL-',
                              'PG  -',
                              'PHST-',
                              'PMC -',
                              'PMID-',
                              'PT  -',
                              'SO  -',
                              'TA  -',
                              'TI  -',
                              'VI  -']

        self.START_PMC = 'PMC -'
        self.START_PMID = 'PMID-'
        self.START_DATE = 'DEP -'
        self.START_TITLE = 'TI  -'
        self.START_ABSTRACT = 'AB  -'
        self.START_LANGUAGE = 'LA  -'
        self.START_PUBLICATION_TYPE = 'PT  -'
        self.START_JOURNAL_TYPE = 'JT  -'
        self.START_DOI = 'SO  -'
        self.START_AUTHOR = 'FAU -'

        self.param_pmc = 'pmc'
        self.param_pmc_pmid = 'pmid'
        self.param_pmc_date = 'data'
        self.param_pmc_title = 'title'
        self.param_pmc_language = 'language'
        self.param_pmc_abstract = 'abstract'
        self.param_pmc_publication_type = 'publication-type'
        self.param_pmc_journal_type = 'journal-type'
        self.param_pmc_doi = 'doi'
        self.param_pmc_author = 'author'

        # ScienceDirect
        self.param_sciencedirect_title = 'T1'
        self.param_sciencedirect_authors = 'AU'
        self.param_sciencedirect_journal = 'JO'
        self.param_sciencedirect_publication_year = 'PY'
        self.param_sciencedirect_volume = 'VL'
        self.param_sciencedirect_first_page = 'SP'
        self.param_sciencedirect_last_page = 'EP'
        self.param_sciencedirect_publication_date = 'DA'
        self.param_sciencedirect_issn = 'SN'
        self.param_sciencedirect_abstract = 'AB'
        self.param_sciencedirect_keywords = 'KW'
        self.param_sciencedirect_doi = 'DO'
        self.param_sciencedirect_sciencedirect_link = 'UR'

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

        if font:
            msg_print = "%s%s%s" % (font, msg_print, self.END)

        if showdate is True:
            _time = time.strftime('%Y-%m-%d %H:%M:%S')
            msg_print = "%s %s" % (_time, msg_print)
            msg_write = "%s %s" % (_time, message)

        print(msg_print)
        if logs:
            for log in logs:
                if log:
                    with open(log, 'a', encoding = 'utf-8') as f:
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
        if path:
            if len(path) > 0 and os.path.exists(path):
                _check = True
        return _check

    def mode_information(self, array1, array2):
        _information = ["%s: %s" % (i, j) for i, j in zip(array1, array2)]
        return " | ".join(_information)

    def read_txt_file(self):
        content = open(self.INPUT_FILE, 'r').readlines()

        collect_unique = {}
        collect_duplicate_doi = {}
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
                    collect[self.xls_col_duplicate_type] = self.xls_val_by_doi
                    collect_duplicate_doi.update({idx: collect})

        collect_papers = {self.XLS_SHEET_UNIQUE: collect_unique,
                          self.XLS_SHEET_DUPLICATES: collect_duplicate_doi}

        return collect_papers

    def read_csv_file(self):

        def check_columns(df, file_name, arr_columns):
            its_ok = True
            for column in arr_columns:
                if column not in df.columns:
                    self.show_print("  Column '%s' don't exist, please check file '%s'" % (column, os.path.basename(file_name)), [self.LOG_FILE], font = self.YELLOW)
                    its_ok = False

            if not its_ok:
                exit()

        use_temporary = False
        _input_file_tmp = ''

        _input_file = self.INPUT_FILE
        if self.TYPE_FILE == self.TYPE_SCOPUS:
            separator = ','
            _col_doi = self.scopus_col_doi
            _col_year = self.scopus_col_year

            arr_columns = [self.scopus_col_authors,
                           self.scopus_col_title,
                           self.scopus_col_abstract,
                           self.scopus_col_year,
                           self.scopus_col_doi,
                           self.scopus_col_document_type,
                           self.scopus_col_language,
                           self.scopus_col_cited_by]
        elif self.TYPE_FILE in [self.TYPE_WOS, self.TYPE_SCIELO]:
            separator = '\t'
            _col_doi = self.wos_col_doi
            _col_year = self.wos_col_year

            arr_columns = [self.wos_col_authors,
                           self.wos_col_title,
                           self.wos_col_abstract,
                           self.wos_col_year,
                           self.wos_col_doi,
                           self.wos_col_document_type,
                           self.wos_col_language,
                           self.wos_col_cited_by]
        elif self.TYPE_FILE == self.TYPE_PUBMED:
            separator = ','
            _col_doi = self.pubmed_col_doi
            _col_year = self.pubmed_col_year

            arr_columns = [self.pubmed_col_authors,
                           self.pubmed_col_title,
                           self.pubmed_col_year,
                           self.pubmed_col_doi]
        elif self.TYPE_FILE == self.TYPE_PUBMED_CENTRAL:
            _input_file = self.read_medline_file(_input_file)
            separator = ','
            _col_doi = self.pmc_col_doi
            _col_year = self.pmc_col_year

            arr_columns = [self.pmc_col_authors,
                           self.pmc_col_title,
                           self.pmc_col_abstract,
                           self.pmc_col_year,
                           self.pmc_col_doi,
                           self.pmc_col_document_type,
                           self.pmc_col_language]
        elif self.TYPE_FILE == self.TYPE_DIMENSIONS:
            _input_file = self.read_dimensions_file(_input_file)
            separator = ','
            _col_doi = self.dimensions_col_doi
            _col_year = self.dimensions_col_year

            arr_columns = [self.dimensions_col_authors,
                           self.dimensions_col_title,
                           self.dimensions_col_abstract,
                           self.dimensions_col_year,
                           self.dimensions_col_doi,
                           self.dimensions_col_cited_by]
        elif self.TYPE_FILE == self.TYPE_GOOGLE_SCHOLAR:
            separator = ','
            _col_doi = self.scholar_col_doi
            _col_year = self.scholar_col_year

            arr_columns = [self.scholar_col_authors,
                           self.scholar_col_title,
                           self.scholar_col_year,
                           self.scholar_col_doi,
                           self.scholar_col_cited_by]
        elif self.TYPE_FILE == self.TYPE_COCHRANE:
            separator = ','
            _col_doi = self.cochrane_col_doi
            _col_year = self.cochrane_col_year

            arr_columns = [self.cochrane_col_authors,
                           self.cochrane_col_title,
                           self.cochrane_col_abstract,
                           self.cochrane_col_year,
                           self.cochrane_col_doi]
        elif self.TYPE_FILE == self.TYPE_EMBASE:
            _input_file = self.read_embase_file(_input_file)
            separator = ','
            _col_doi = self.embase_col_doi
            _col_year = self.embase_col_year

            arr_columns = [self.embase_col_authors,
                           self.embase_col_title,
                           self.embase_col_abstract,
                           self.embase_col_year,
                           self.embase_col_doi,
                           self.embase_col_document_type,
                           self.embase_col_language]

        elif self.TYPE_FILE == self.TYPE_SCIENCEDIRECT:
            _input_file = self.read_sciencedirect_file(_input_file)
            separator = '|'
            _col_doi = self.sciencedirect_col_doi
            _col_year = self.sciencedirect_col_year

            arr_columns = [self.sciencedirect_col_authors,
                           self.sciencedirect_col_title,
                           self.sciencedirect_col_abstract,
                           self.sciencedirect_col_year,
                           self.sciencedirect_col_doi]
        elif self.TYPE_FILE == self.TYPE_IEEE:
            separator = ','
            _col_doi = self.ieee_col_doi
            _col_year = self.ieee_col_year

            arr_columns = [self.ieee_col_authors,
                           self.ieee_col_title,
                           self.ieee_col_abstract,
                           self.ieee_col_year,
                           self.ieee_col_doi]
        elif self.TYPE_FILE == self.TYPE_BVS:
            separator = ','
            _col_doi = self.bvs_col_doi
            _col_year = self.bvs_col_year

            arr_columns = [self.bvs_col_authors,
                           self.bvs_col_title,
                           self.bvs_col_abstract,
                           self.bvs_col_year,
                           self.bvs_col_doi,
                           self.bvs_col_document_type,
                           self.bvs_col_language]
        elif self.TYPE_FILE == self.TYPE_CAB:
            separator = ','
            _col_doi = self.cab_col_doi
            _col_year = self.cab_col_year

            arr_columns = [self.cab_col_authors,
                           self.cab_col_title,
                           self.cab_col_abstract,
                           self.cab_col_year,
                           self.cab_col_doi,
                           self.cab_col_language]

        if self.TYPE_FILE in [self.TYPE_BVS, self.TYPE_CAB]:
            df = pd.read_csv(filepath_or_buffer = _input_file, sep = separator, header = 0, index_col = False, on_bad_lines = 'skip')
        else:
            df = pd.read_csv(filepath_or_buffer = _input_file, sep = separator, header = 0, index_col = False) # low_memory = False

        # df = df.where(pd.notnull(df), '') # None
        df = df.replace({np.nan: ''}) # None
        df.columns = df.columns.str.strip()
        # print(df)

        # Check columns
        check_columns(df, _input_file, arr_columns)

        # Get DOIs
        collect_unique_doi = {}
        collect_duplicate_doi = {}
        collect_without_doi = {}
        nr_doi = []
        for idx, row in df.iterrows():
            flag_unique = False
            flag_duplicate_doi = False
            flag_without_doi = False

            doi = row[_col_doi]
            doi = doi.strip()
            if doi:
                doi = doi.lower()
                doi = doi[:-1] if doi.endswith('.') else doi

                if 'doi.org' in doi:
                    doi = doi.split('.org/')[1]

                pattern = re.compile(r'^10\.')
                if pattern.match(doi):
                    if doi not in nr_doi:
                        nr_doi.append(doi)
                        flag_unique = True
                    else:
                        flag_duplicate_doi = True
                else:
                    doi = ''
                    flag_without_doi = True
            else:
                flag_without_doi = True

            year = row[_col_year]
            if year:
                year = str(year).strip()

            pattern = re.compile(r'^\d{4}$')
            if pattern.match(year):
                year = int(year)
            else:
                year = None

            collect = {}
            if self.TYPE_FILE == self.TYPE_SCOPUS:
                collect[self.xls_col_authors] = row[self.scopus_col_authors].strip() if row[self.scopus_col_authors] else row[self.scopus_col_authors]
                collect[self.xls_col_title] = row[self.scopus_col_title].strip() if row[self.scopus_col_title] else row[self.scopus_col_title]
                collect[self.xls_col_abstract] = row[self.scopus_col_abstract].strip() if row[self.scopus_col_abstract] else row[self.scopus_col_abstract]
                collect[self.xls_col_year] = year
                collect[self.xls_col_doi] = doi
                collect[self.xls_col_document_type] = row[self.scopus_col_document_type].strip() if row[self.scopus_col_document_type] else row[self.scopus_col_document_type]
                collect[self.xls_col_language] = row[self.scopus_col_language].strip() if row[self.scopus_col_language] else row[self.scopus_col_language]
                collect[self.xls_col_cited_by] = row[self.scopus_col_cited_by] if row[self.scopus_col_cited_by] else 0
            elif self.TYPE_FILE in [self.TYPE_WOS, self.TYPE_SCIELO]:
                document_type = row[self.wos_col_document_type]
                document_type = self.format_publication_type(document_type)
                collect[self.xls_col_authors] = row[self.wos_col_authors].strip() if row[self.wos_col_authors] else row[self.wos_col_authors]
                collect[self.xls_col_title] = row[self.wos_col_title].strip() if row[self.wos_col_title] else row[self.wos_col_title]
                collect[self.xls_col_abstract] = row[self.wos_col_abstract].strip() if row[self.wos_col_abstract] else row[self.wos_col_abstract]
                collect[self.xls_col_year] = year
                collect[self.xls_col_doi] = doi
                collect[self.xls_col_document_type] = document_type if document_type else document_type
                collect[self.xls_col_language] = row[self.wos_col_language].strip() if row[self.wos_col_language] else row[self.wos_col_language]
                collect[self.xls_col_cited_by] = row[self.wos_col_cited_by] if row[self.wos_col_cited_by] else row[self.wos_col_cited_by]
            elif self.TYPE_FILE == self.TYPE_PUBMED:
                collect[self.xls_col_authors] = row[self.pubmed_col_authors].strip() if row[self.pubmed_col_authors] else row[self.pubmed_col_authors]
                collect[self.xls_col_title] = row[self.pubmed_col_title].strip() if row[self.pubmed_col_title] else row[self.pubmed_col_title]
                collect[self.xls_col_abstract] = None
                collect[self.xls_col_year] = year
                collect[self.xls_col_doi] = doi
                collect[self.xls_col_document_type] = None
                collect[self.xls_col_language] = None
                collect[self.xls_col_cited_by] = None
            elif self.TYPE_FILE == self.TYPE_PUBMED_CENTRAL:
                collect[self.xls_col_authors] = row[self.pmc_col_authors].strip() if row[self.pmc_col_authors] else row[self.pmc_col_authors]
                collect[self.xls_col_title] = row[self.pmc_col_title].strip() if row[self.pmc_col_title] else row[self.pmc_col_title]
                collect[self.xls_col_abstract] = row[self.pmc_col_abstract].strip() if row[self.pmc_col_abstract] else row[self.pmc_col_abstract]
                collect[self.xls_col_year] = year
                collect[self.xls_col_doi] = doi
                collect[self.xls_col_document_type] = row[self.pmc_col_document_type].strip() if row[self.pmc_col_document_type] else row[self.pmc_col_document_type]
                collect[self.xls_col_language] = row[self.pmc_col_language].strip() if row[self.pmc_col_language] else row[self.pmc_col_language]
                collect[self.xls_col_cited_by] = None
            elif self.TYPE_FILE == self.TYPE_DIMENSIONS:
                collect[self.xls_col_authors] = row[self.dimensions_col_authors].strip() if row[self.dimensions_col_authors] else row[self.dimensions_col_authors]
                collect[self.xls_col_title] = row[self.dimensions_col_title].strip() if row[self.dimensions_col_title] else row[self.dimensions_col_title]
                collect[self.xls_col_abstract] = row[self.dimensions_col_abstract].strip() if row[self.dimensions_col_abstract] else row[self.dimensions_col_abstract]
                collect[self.xls_col_year] = year
                collect[self.xls_col_doi] = doi
                collect[self.xls_col_document_type] = None
                collect[self.xls_col_language] = None
                collect[self.xls_col_cited_by] = row[self.dimensions_col_cited_by] if row[self.dimensions_col_cited_by] else row[self.dimensions_col_cited_by]
            elif self.TYPE_FILE == self.TYPE_GOOGLE_SCHOLAR:
                collect[self.xls_col_authors] = row[self.scholar_col_authors].strip() if row[self.scholar_col_authors] else row[self.scholar_col_authors]
                collect[self.xls_col_title] = row[self.scholar_col_title].strip() if row[self.scholar_col_title] else row[self.scholar_col_title]
                collect[self.xls_col_abstract] = None
                collect[self.xls_col_year] = year
                collect[self.xls_col_doi] = doi
                collect[self.xls_col_document_type] = None
                collect[self.xls_col_language] = None
                collect[self.xls_col_cited_by] = row[self.scholar_col_cited_by] if row[self.scholar_col_cited_by] else row[self.scholar_col_cited_by]
            elif self.TYPE_FILE == self.TYPE_COCHRANE:
                collect[self.xls_col_authors] = row[self.cochrane_col_authors].strip() if row[self.cochrane_col_authors] else row[self.cochrane_col_authors]
                collect[self.xls_col_title] = row[self.cochrane_col_title].strip() if row[self.cochrane_col_title] else row[self.cochrane_col_title]
                collect[self.xls_col_abstract] = row[self.cochrane_col_abstract].strip() if row[self.cochrane_col_abstract] else row[self.cochrane_col_abstract]
                collect[self.xls_col_year] = year
                collect[self.xls_col_doi] = doi
                collect[self.xls_col_document_type] = None
                collect[self.xls_col_language] = None
                collect[self.xls_col_cited_by] = None
            elif self.TYPE_FILE == self.TYPE_EMBASE:
                collect[self.xls_col_authors] = row[self.embase_col_authors].strip() if row[self.embase_col_authors] else row[self.embase_col_authors]
                collect[self.xls_col_title] = row[self.embase_col_title].strip() if row[self.embase_col_title] else row[self.embase_col_title]
                collect[self.xls_col_abstract] = row[self.embase_col_abstract].strip() if row[self.embase_col_abstract] else row[self.embase_col_abstract]
                collect[self.xls_col_year] = year
                collect[self.xls_col_doi] = doi
                collect[self.xls_col_document_type] = row[self.embase_col_document_type].strip() if row[self.embase_col_document_type] else row[self.embase_col_document_type]
                collect[self.xls_col_language] = row[self.embase_col_language].strip() if row[self.embase_col_language] else row[self.embase_col_language]
                collect[self.xls_col_cited_by] = None
            elif self.TYPE_FILE == self.TYPE_SCIENCEDIRECT:
                collect[self.xls_col_authors] = row[self.sciencedirect_col_authors].strip() if row[self.sciencedirect_col_authors] else row[self.sciencedirect_col_authors]
                collect[self.xls_col_title] = row[self.sciencedirect_col_title].strip() if row[self.sciencedirect_col_title] else row[self.sciencedirect_col_title]
                collect[self.xls_col_abstract] = row[self.sciencedirect_col_abstract].strip() if row[self.sciencedirect_col_abstract] else row[self.sciencedirect_col_abstract]
                collect[self.xls_col_year] = year
                collect[self.xls_col_doi] = doi
                collect[self.xls_col_document_type] = None
                collect[self.xls_col_language] = None
                collect[self.xls_col_cited_by] = None
            elif self.TYPE_FILE == self.TYPE_IEEE:
                collect[self.xls_col_authors] = row[self.ieee_col_authors].strip() if row[self.ieee_col_authors] else row[self.ieee_col_authors]
                collect[self.xls_col_title] = row[self.ieee_col_title].strip() if row[self.ieee_col_title] else row[self.ieee_col_title]
                collect[self.xls_col_abstract] = row[self.ieee_col_abstract].strip() if row[self.ieee_col_abstract] else row[self.ieee_col_abstract]
                collect[self.xls_col_year] = year
                collect[self.xls_col_doi] = doi
                collect[self.xls_col_document_type] = None
                collect[self.xls_col_language] = None
                collect[self.xls_col_cited_by] = None
            elif self.TYPE_FILE == self.TYPE_BVS:
                language = row[self.bvs_col_language]
                language = self.get_language(language)
                document_type = row[self.bvs_col_document_type].capitalize()
                collect[self.xls_col_authors] = row[self.bvs_col_authors].strip() if row[self.bvs_col_authors] else row[self.bvs_col_authors]
                collect[self.xls_col_title] = row[self.bvs_col_title].strip() if row[self.bvs_col_title] else row[self.bvs_col_title]
                collect[self.xls_col_abstract] = row[self.bvs_col_abstract].strip() if row[self.bvs_col_abstract] else row[self.bvs_col_abstract]
                collect[self.xls_col_year] = year
                collect[self.xls_col_doi] = doi
                collect[self.xls_col_document_type] = document_type if document_type else document_type
                collect[self.xls_col_language] = language if language else language
                collect[self.xls_col_cited_by] = None
            elif self.TYPE_FILE == self.TYPE_CAB:
                collect[self.xls_col_authors] = row[self.cab_col_authors].strip() if row[self.cab_col_authors] else row[self.cab_col_authors]
                collect[self.xls_col_title] = row[self.cab_col_title].strip() if row[self.cab_col_title] else row[self.cab_col_title]
                collect[self.xls_col_abstract] = row[self.cab_col_abstract].strip() if row[self.cab_col_abstract] else row[self.cab_col_abstract]
                collect[self.xls_col_year] = year
                collect[self.xls_col_doi] = doi
                collect[self.xls_col_document_type] = None
                collect[self.xls_col_language] = row[self.cab_col_language].strip() if row[self.cab_col_language] else row[self.cab_col_language]
                collect[self.xls_col_cited_by] = None

            if flag_unique:
                collect_unique_doi.update({idx + 1: collect})
            if flag_duplicate_doi:
                collect[self.xls_col_duplicate_type] = self.xls_val_by_doi
                collect_duplicate_doi.update({idx + 1: collect})
            if flag_without_doi:
                collect_without_doi.update({idx + 1: collect})

        # Get titles
        collect_unique = {}
        collect_duplicate_title = {}
        nr_title = []
        index = 1
        for idx, row in collect_unique_doi.items():
            flag_unique = False

            title = row[self.xls_col_title]
            if title:
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
                row[self.xls_col_duplicate_type] = self.xls_val_by_title
                collect_duplicate_title.update({idx: row})

        collect_duplicate = {}
        collect_duplicate = collect_duplicate_doi.copy()
        collect_duplicate.update(collect_duplicate_title)
        collect_duplicate = {item[0]: item[1] for item in sorted(collect_duplicate.items())}

        collect_papers = {self.XLS_SHEET_UNIQUE: collect_unique,
                          self.XLS_SHEET_WITHOUT_DOI: collect_without_doi,
                          self.XLS_SHEET_DUPLICATES: collect_duplicate}

        return collect_papers

    def save_summary_xls(self, data_paper):

        def create_sheet(oworkbook, sheet_type, dictionary, styles_title, styles_rows):
            if self.TYPE_FILE == self.TYPE_TXT:
                _xls_columns = self.xls_columns_txt.copy()
            else:
                _xls_columns = self.xls_columns_csv.copy()

            if sheet_type == self.XLS_SHEET_DUPLICATES:
                _xls_columns.append(self.xls_col_duplicate_type)

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
                if sheet_type == self.XLS_SHEET_DUPLICATES:
                    worksheet.set_column(first_col = 2, last_col = 2, width = 19) # Column C:C
            else:
                worksheet.set_column(first_col = 0, last_col = 0, width = 7)  # Column A:A
                worksheet.set_column(first_col = 1, last_col = 1, width = 30) # Column B:B
                worksheet.set_column(first_col = 2, last_col = 2, width = 33) # Column C:C
                worksheet.set_column(first_col = 3, last_col = 3, width = 8)  # Column D:D
                worksheet.set_column(first_col = 4, last_col = 4, width = 30) # Column E:E
                worksheet.set_column(first_col = 5, last_col = 5, width = 18) # Column F:F
                worksheet.set_column(first_col = 6, last_col = 6, width = 12) # Column G:G
                worksheet.set_column(first_col = 7, last_col = 7, width = 11) # Column H:H
                worksheet.set_column(first_col = 8, last_col = 8, width = 18) # Column I:I
                if sheet_type == self.XLS_SHEET_DUPLICATES:
                    worksheet.set_column(first_col = 9, last_col = 9, width = 17) # Column J:J

            icol = 0
            for irow, (index, item) in enumerate(dictionary.items(), start = 1):
                col_doi = item[self.xls_col_doi]
                if sheet_type == self.XLS_SHEET_DUPLICATES:
                    duplicate_type = item[self.xls_col_duplicate_type]

                if self.TYPE_FILE == self.TYPE_TXT:
                    worksheet.write(irow, icol + 0, index, styles_rows)
                    worksheet.write(irow, icol + 1, col_doi, styles_rows)
                    if sheet_type == self.XLS_SHEET_DUPLICATES:
                        worksheet.write(irow, icol + 2, duplicate_type, styles_rows)
                else:
                    worksheet.write(irow, icol + 0, index, styles_rows)
                    worksheet.write(irow, icol + 1, item[self.xls_col_title], styles_rows)
                    worksheet.write(irow, icol + 2, item[self.xls_col_abstract], styles_rows)
                    worksheet.write(irow, icol + 3, item[self.xls_col_year], styles_rows)
                    worksheet.write(irow, icol + 4, col_doi, styles_rows)
                    worksheet.write(irow, icol + 5, item[self.xls_col_document_type], styles_rows)
                    worksheet.write(irow, icol + 6, item[self.xls_col_language], styles_rows)
                    worksheet.write(irow, icol + 7, item[self.xls_col_cited_by], styles_rows)
                    worksheet.write(irow, icol + 8, item[self.xls_col_authors], styles_rows)
                    if sheet_type == self.XLS_SHEET_DUPLICATES:
                        worksheet.write(irow, icol + 9, duplicate_type, styles_rows)

        workbook = xlsxwriter.Workbook(self.XLS_FILE)

        # Styles
        cell_format_title = workbook.add_format({'bold': True,
                                                 'font_color': 'white',
                                                 'bg_color': 'black',
                                                 'align': 'center',
                                                 'valign': 'vcenter'})
        cell_format_row = workbook.add_format({'text_wrap': True, 'valign': 'top'})

        create_sheet(workbook, self.XLS_SHEET_UNIQUE, data_paper[self.XLS_SHEET_UNIQUE], cell_format_title, cell_format_row)
        if self.TYPE_FILE != self.TYPE_TXT:
            create_sheet(workbook, self.XLS_SHEET_WITHOUT_DOI, data_paper[self.XLS_SHEET_WITHOUT_DOI], cell_format_title, cell_format_row)
        create_sheet(workbook, self.XLS_SHEET_DUPLICATES, data_paper[self.XLS_SHEET_DUPLICATES], cell_format_title, cell_format_row)

        workbook.close()

    def get_language(self, code):
        # https://en.wikipedia.org/wiki/List_of_ISO_639_language_codes
        hash_data = {
            'ab': 'Abkhazian',
            'aa': 'Afar',
            'af': 'Afrikaans',
            'ak': 'Akan',
            'sq': 'Albanian',
            'am': 'Amharic',
            'ar': 'Arabic',
            'an': 'Aragonese',
            'hy': 'Armenian',
            'as': 'Assamese',
            'av': 'Avaric',
            'ae': 'Avestan',
            'ay': 'Aymara',
            'az': 'Azerbaijani',
            'bm': 'Bambara',
            'ba': 'Bashkir',
            'eu': 'Basque',
            'be': 'Belarusian',
            'bn': 'Bengali',
            'bi': 'Bislama',
            'bs': 'Bosnian',
            'br': 'Breton',
            'bg': 'Bulgarian',
            'my': 'Burmese',
            'ca': 'Catalan, Valencian',
            'km': 'Central Khmer',
            'ch': 'Chamorro',
            'ce': 'Chechen',
            'ny': 'Chichewa, Chewa, Nyanja',
            'zh': 'Chinese',
            'cu': 'Church Slavonic, Old Slavonic, Old Church Slavonic',
            'cv': 'Chuvash',
            'kw': 'Cornish',
            'co': 'Corsican',
            'cr': 'Cree',
            'hr': 'Croatian',
            'cs': 'Czech',
            'da': 'Danish',
            'dv': 'Divehi, Dhivehi, Maldivian',
            'nl': 'Dutch, Flemish',
            'dz': 'Dzongkha',
            'en': 'English',
            'eo': 'Esperanto',
            'et': 'Estonian',
            'ee': 'Ewe',
            'fo': 'Faroese',
            'fj': 'Fijian',
            'fi': 'Finnish',
            'fr': 'French',
            'ff': 'Fulah',
            'gd': 'Gaelic, Scottish Gaelic',
            'gl': 'Galician',
            'lg': 'Ganda',
            'ka': 'Georgian',
            'de': 'German',
            'el': 'Greek, Modern (1453–)',
            'gn': 'Guarani',
            'gu': 'Gujarati',
            'ht': 'Haitian, Haitian Creole',
            'ha': 'Hausa',
            'he': 'Hebrew',
            'hz': 'Herero',
            'hi': 'Hindi',
            'ho': 'Hiri Motu',
            'hu': 'Hungarian',
            'is': 'Icelandic',
            'io': 'Ido',
            'ig': 'Igbo',
            'id': 'Indonesian',
            'ia': 'Interlingua (International Auxiliary Language Association)',
            'ie': 'Interlingue, Occidental',
            'iu': 'Inuktitut',
            'ik': 'Inupiaq',
            'ga': 'Irish',
            'it': 'Italian',
            'ja': 'Japanese',
            'jv': 'Javanese',
            'kl': 'Kalaallisut, Greenlandic',
            'kn': 'Kannada',
            'kr': 'Kanuri',
            'ks': 'Kashmiri',
            'kk': 'Kazakh',
            'ki': 'Kikuyu, Gikuyu',
            'rw': 'Kinyarwanda',
            'kv': 'Komi',
            'kg': 'Kongo',
            'ko': 'Korean',
            'kj': 'Kuanyama, Kwanyama',
            'ku': 'Kurdish',
            'ky': 'Kyrgyz, Kirghiz',
            'lo': 'Lao',
            'la': 'Latin',
            'lv': 'Latvian',
            'li': 'Limburgan, Limburger, Limburgish',
            'ln': 'Lingala',
            'lt': 'Lithuanian',
            'lu': 'Luba-Katanga',
            'lb': 'Luxembourgish, Letzeburgesch',
            'mk': 'Macedonian',
            'mg': 'Malagasy',
            'ms': 'Malay',
            'ml': 'Malayalam',
            'mt': 'Maltese',
            'gv': 'Manx',
            'mi': 'Maori',
            'mr': 'Marathi',
            'mh': 'Marshallese',
            'mn': 'Mongolian',
            'na': 'Nauru',
            'nv': 'Navajo, Navaho',
            'ng': 'Ndonga',
            'ne': 'Nepali',
            'nd': 'North Ndebele',
            'se': 'Northern Sami',
            'no': 'Norwegian',
            'nb': 'Norwegian Bokmål',
            'nn': 'Norwegian Nynorsk',
            'oc': 'Occitan',
            'oj': 'Ojibwa',
            'or': 'Oriya',
            'om': 'Oromo',
            'os': 'Ossetian, Ossetic',
            'pi': 'Pali',
            'ps': 'Pashto, Pushto',
            'fa': 'Persian',
            'pl': 'Polish',
            'pt': 'Portuguese',
            'pa': 'Punjabi, Panjabi',
            'qu': 'Quechua',
            'ro': 'Romanian, Moldavian, Moldovan',
            'rm': 'Romansh',
            'rn': 'Rundi',
            'ru': 'Russian',
            'sm': 'Samoan',
            'sg': 'Sango',
            'sa': 'Sanskrit',
            'sc': 'Sardinian',
            'sr': 'Serbian',
            'sn': 'Shona',
            'ii': 'Sichuan Yi, Nuosu',
            'sd': 'Sindhi',
            'si': 'Sinhala, Sinhalese',
            'sk': 'Slovak',
            'sl': 'Slovenian',
            'so': 'Somali',
            'nr': 'South Ndebele',
            'st': 'Southern Sotho',
            'es': 'Spanish, Castilian',
            'su': 'Sundanese',
            'sw': 'Swahili',
            'ss': 'Swati',
            'sv': 'Swedish',
            'tl': 'Tagalog',
            'ty': 'Tahitian',
            'tg': 'Tajik',
            'ta': 'Tamil',
            'tt': 'Tatar',
            'te': 'Telugu',
            'th': 'Thai',
            'bo': 'Tibetan',
            'ti': 'Tigrinya',
            'to': 'Tonga (Tonga Islands)',
            'ts': 'Tsonga',
            'tn': 'Tswana',
            'tr': 'Turkish',
            'tk': 'Turkmen',
            'tw': 'Twi',
            'ug': 'Uighur, Uyghur',
            'uk': 'Ukrainian',
            'ur': 'Urdu',
            'uz': 'Uzbek',
            've': 'Venda',
            'vi': 'Vietnamese',
            'vo': 'Volapük',
            'wa': 'Walloon',
            'cy': 'Welsh',
            'fy': 'Western Frisian',
            'wo': 'Wolof',
            'xh': 'Xhosa',
            'yi': 'Yiddish',
            'yo': 'Yoruba',
            'za': 'Zhuang, Chuang',
            'zu': 'Zulu',
            'abk': 'Abkhazian',
            'aar': 'Afar',
            'afr': 'Afrikaans',
            'aka': 'Akan',
            'sqi': 'Albanian',
            'amh': 'Amharic',
            'ara': 'Arabic',
            'arg': 'Aragonese',
            'hye': 'Armenian',
            'asm': 'Assamese',
            'ava': 'Avaric',
            'ave': 'Avestan',
            'aym': 'Aymara',
            'aze': 'Azerbaijani',
            'bam': 'Bambara',
            'bak': 'Bashkir',
            'eus': 'Basque',
            'bel': 'Belarusian',
            'ben': 'Bengali',
            'bis': 'Bislama',
            'bos': 'Bosnian',
            'bre': 'Breton',
            'bul': 'Bulgarian',
            'mya': 'Burmese',
            'cat': 'Catalan, Valencian',
            'khm': 'Central Khmer',
            'cha': 'Chamorro',
            'che': 'Chechen',
            'nya': 'Chichewa, Chewa, Nyanja',
            'zho': 'Chinese',
            'chu': 'Church Slavonic, Old Slavonic, Old Church Slavonic',
            'chv': 'Chuvash',
            'cor': 'Cornish',
            'cos': 'Corsican',
            'cre': 'Cree',
            'hrv': 'Croatian',
            'ces': 'Czech',
            'dan': 'Danish',
            'div': 'Divehi, Dhivehi, Maldivian',
            'nld': 'Dutch, Flemish',
            'dzo': 'Dzongkha',
            'eng': 'English',
            'epo': 'Esperanto',
            'est': 'Estonian',
            'ewe': 'Ewe',
            'fao': 'Faroese',
            'fij': 'Fijian',
            'fin': 'Finnish',
            'fra': 'French',
            'ful': 'Fulah',
            'gla': 'Gaelic, Scottish Gaelic',
            'glg': 'Galician',
            'lug': 'Ganda',
            'kat': 'Georgian',
            'deu': 'German',
            'ell': 'Greek, Modern (1453–)',
            'grn': 'Guarani',
            'guj': 'Gujarati',
            'hat': 'Haitian, Haitian Creole',
            'hau': 'Hausa',
            'heb': 'Hebrew',
            'her': 'Herero',
            'hin': 'Hindi',
            'hmo': 'Hiri Motu',
            'hun': 'Hungarian',
            'isl': 'Icelandic',
            'ido': 'Ido',
            'ibo': 'Igbo',
            'ind': 'Indonesian',
            'ina': 'Interlingua (International Auxiliary Language Association)',
            'ile': 'Interlingue, Occidental',
            'iku': 'Inuktitut',
            'ipk': 'Inupiaq',
            'gle': 'Irish',
            'ita': 'Italian',
            'jpn': 'Japanese',
            'jav': 'Javanese',
            'kal': 'Kalaallisut, Greenlandic',
            'kan': 'Kannada',
            'kau': 'Kanuri',
            'kas': 'Kashmiri',
            'kaz': 'Kazakh',
            'kik': 'Kikuyu, Gikuyu',
            'kin': 'Kinyarwanda',
            'kom': 'Komi',
            'kon': 'Kongo',
            'kor': 'Korean',
            'kua': 'Kuanyama, Kwanyama',
            'kur': 'Kurdish',
            'kir': 'Kyrgyz, Kirghiz',
            'lao': 'Lao',
            'lat': 'Latin',
            'lav': 'Latvian',
            'lim': 'Limburgan, Limburger, Limburgish',
            'lin': 'Lingala',
            'lit': 'Lithuanian',
            'lub': 'Luba-Katanga',
            'ltz': 'Luxembourgish, Letzeburgesch',
            'mkd': 'Macedonian',
            'mlg': 'Malagasy',
            'msa': 'Malay',
            'mal': 'Malayalam',
            'mlt': 'Maltese',
            'glv': 'Manx',
            'mri': 'Maori',
            'mar': 'Marathi',
            'mah': 'Marshallese',
            'mon': 'Mongolian',
            'nau': 'Nauru',
            'nav': 'Navajo, Navaho',
            'ndo': 'Ndonga',
            'nep': 'Nepali',
            'nde': 'North Ndebele',
            'sme': 'Northern Sami',
            'nor': 'Norwegian',
            'nob': 'Norwegian Bokmål',
            'nno': 'Norwegian Nynorsk',
            'oci': 'Occitan',
            'oji': 'Ojibwa',
            'ori': 'Oriya',
            'orm': 'Oromo',
            'oss': 'Ossetian, Ossetic',
            'pli': 'Pali',
            'pus': 'Pashto, Pushto',
            'fas': 'Persian',
            'pol': 'Polish',
            'por': 'Portuguese',
            'pan': 'Punjabi, Panjabi',
            'que': 'Quechua',
            'ron': 'Romanian, Moldavian, Moldovan',
            'roh': 'Romansh',
            'run': 'Rundi',
            'rus': 'Russian',
            'smo': 'Samoan',
            'sag': 'Sango',
            'san': 'Sanskrit',
            'srd': 'Sardinian',
            'srp': 'Serbian',
            'sna': 'Shona',
            'iii': 'Sichuan Yi, Nuosu',
            'snd': 'Sindhi',
            'sin': 'Sinhala, Sinhalese',
            'slk': 'Slovak',
            'slv': 'Slovenian',
            'som': 'Somali',
            'nbl': 'South Ndebele',
            'sot': 'Southern Sotho',
            'spa': 'Spanish, Castilian',
            'sun': 'Sundanese',
            'swa': 'Swahili',
            'ssw': 'Swati',
            'swe': 'Swedish',
            'tgl': 'Tagalog',
            'tah': 'Tahitian',
            'tgk': 'Tajik',
            'tam': 'Tamil',
            'tat': 'Tatar',
            'tel': 'Telugu',
            'tha': 'Thai',
            'bod': 'Tibetan',
            'tir': 'Tigrinya',
            'ton': 'Tonga (Tonga Islands)',
            'tso': 'Tsonga',
            'tsn': 'Tswana',
            'tur': 'Turkish',
            'tuk': 'Turkmen',
            'twi': 'Twi',
            'uig': 'Uighur, Uyghur',
            'ukr': 'Ukrainian',
            'urd': 'Urdu',
            'uzb': 'Uzbek',
            'ven': 'Venda',
            'vie': 'Vietnamese',
            'vol': 'Volapük',
            'wln': 'Walloon',
            'cym': 'Welsh',
            'fry': 'Western Frisian',
            'wol': 'Wolof',
            'xho': 'Xhosa',
            'yid': 'Yiddish',
            'yor': 'Yoruba',
            'zha': 'Zhuang, Chuang',
            'zul': 'Zulu',
            'alb': 'Albanian',
            'arm': 'Armenian',
            'baq': 'Basque',
            'bur': 'Burmese',
            'chi': 'Chinese',
            'cze': 'Czech',
            'dut': 'Dutch, Flemish',
            'fre': 'French',
            'geo': 'Georgian',
            'ger': 'German',
            'gre': 'Greek, Modern (1453–)',
            'ice': 'Icelandic',
            'mac': 'Macedonian',
            'may': 'Malay',
            'mao': 'Maori',
            'per': 'Persian',
            'rum': 'Romanian, Moldavian, Moldovan',
            'slo': 'Slovak',
            'tib': 'Tibetan',
            'wel': 'Welsh'
        }

        r = 'Unknown'
        if code in hash_data:
            r = hash_data[code]

        return r

    def remove_endpoint(self, text):
        _text = text.strip()

        while(_text[-1] == '.'):
            _text = _text[0:len(_text) - 1]
            _text = _text.strip()

        return _text

    def block_continue(self, text):
        _continue = True
        for _start in self.MEDLINE_START:
            if text.startswith(_start):
                _continue = False
                break
        return _continue

    def get_data(self, text, array, start_param):
        if text.startswith(start_param):
            _line = text.replace(start_param, '').strip()
            array.append(_line)
            # continue

    def format_publication_type(self, publication_type):
        replacements = {'research-article': 'Article',
                        'review-article': 'Review'}

        if publication_type in replacements:
            publication = replacements[publication_type]
        elif '-' in publication_type:
            publication = publication_type.replace('-', ' ').title()
        else:
            publication = publication_type.title()

        return publication

    def read_medline_file(self, file):

        def rename_publication_type(text):
            doc_type = None
            if text == 'Journal Article':
                doc_type = 'Article'
            elif text == 'Journal Article Case Report':
                doc_type = 'Case Report'
            elif text == 'Journal Article Editorial':
                doc_type = 'Editorial'
            elif text == 'Journal Article Letter':
                doc_type = 'Letter'
            elif text == 'Journal Article News':
                doc_type = 'News'
            elif text == 'Journal Article Review':
                doc_type = 'Review'
            else:
                doc_type = text
            return doc_type

        medline_data = {}
        with open(file, 'r', encoding = 'utf8') as fr:
            item_dict = {self.param_pmc: None,
                         self.param_pmc_pmid: None,
                         self.param_pmc_date: None,
                         self.param_pmc_title: None,
                         self.param_pmc_language: None,
                         self.param_pmc_abstract: None,
                         self.param_pmc_publication_type: None,
                         self.param_pmc_journal_type: None,
                         self.param_pmc_doi: None,
                         self.param_pmc_author: None}
            index = 1

            flag_start = False
            flag_title = False
            flag_abstract = False
            flag_doi = False

            arr_pmc = []
            arr_pmid = []
            arr_language = []
            arr_journal_type = []
            arr_publication_type = []
            arr_date = []
            arr_title = []
            arr_abstract = []
            arr_doi = []
            arr_author = []

            for line in fr:
                line = line.strip()
                if line:
                    # PMC
                    if line.startswith(self.START_PMC):
                        # Check
                        if arr_pmc:
                            _item_dict = item_dict.copy()
                            _item_dict.update({self.param_pmc: arr_pmc})
                            _item_dict.update({self.param_pmc_pmid: arr_pmid})
                            _item_dict.update({self.param_pmc_language: arr_language})
                            _item_dict.update({self.param_pmc_journal_type: arr_journal_type})
                            _item_dict.update({self.param_pmc_publication_type: arr_publication_type})
                            _item_dict.update({self.param_pmc_date: arr_date})
                            _item_dict.update({self.param_pmc_title: arr_title})
                            _item_dict.update({self.param_pmc_abstract: arr_abstract})
                            _item_dict.update({self.param_pmc_doi: arr_doi})
                            _item_dict.update({self.param_pmc_author: arr_author})
                            medline_data.update({index: _item_dict})
                            index += 1

                            flag_start = False
                            flag_title = False
                            flag_abstract = False
                            flag_doi = False

                            arr_pmc = []
                            arr_pmid = []
                            arr_language = []
                            arr_journal_type = []
                            arr_publication_type = []
                            arr_date = []
                            arr_title = []
                            arr_abstract = []
                            arr_doi = []
                            arr_author = []

                        flag_start = True
                        _line = line.replace(self.START_PMC, '').strip()
                        arr_pmc.append(_line)
                        continue

                    if flag_start:
                        self.get_data(line, arr_pmid, self.START_PMID)
                        self.get_data(line, arr_language, self.START_LANGUAGE)
                        self.get_data(line, arr_journal_type, self.START_JOURNAL_TYPE)
                        self.get_data(line, arr_publication_type, self.START_PUBLICATION_TYPE)
                        self.get_data(line, arr_date, self.START_DATE)
                        self.get_data(line, arr_author, self.START_AUTHOR)

                        # Title
                        if line.startswith(self.START_TITLE):
                            flag_title = True
                            _line = line.replace(self.START_TITLE, '').strip()
                            arr_title.append(_line)
                            continue
                        if flag_title:
                            if self.block_continue(line):
                                arr_title.append(line)
                                continue
                            else:
                                flag_title = False

                        # Abstract
                        if line.startswith(self.START_ABSTRACT):
                            flag_abstract = True
                            _line = line.replace(self.START_ABSTRACT, '').strip()
                            arr_abstract.append(_line)
                            continue
                        if flag_abstract:
                            if self.block_continue(line):
                                arr_abstract.append(line)
                                continue
                            else:
                                flag_abstract = False

                        # DOI
                        if line.startswith(self.START_DOI):
                            flag_doi = True
                            _line = line.replace(self.START_DOI, '').strip()
                            arr_doi.append(_line)
                            continue
                        if flag_doi:
                            if self.block_continue(line):
                                arr_doi.append(line)
                                continue
                            else:
                                flag_doi = False

            if arr_pmc:
                _item_dict = item_dict.copy()
                _item_dict.update({self.param_pmc: arr_pmc})
                _item_dict.update({self.param_pmc_pmid: arr_pmid})
                _item_dict.update({self.param_pmc_language: arr_language})
                _item_dict.update({self.param_pmc_journal_type: arr_journal_type})
                _item_dict.update({self.param_pmc_publication_type: arr_publication_type})
                _item_dict.update({self.param_pmc_date: arr_date})
                _item_dict.update({self.param_pmc_title: arr_title})
                _item_dict.update({self.param_pmc_abstract: arr_abstract})
                _item_dict.update({self.param_pmc_doi: arr_doi})
                _item_dict.update({self.param_pmc_author: arr_author})
                medline_data.update({index: _item_dict})
        fr.close()

        for index, item in medline_data.items():
            _publication_type = rename_publication_type(' '.join(item[self.param_pmc_publication_type]))
            item.update({self.param_pmc: ' '.join(item[self.param_pmc])})
            item.update({self.param_pmc_pmid: ' '.join(item[self.param_pmc_pmid])})
            item.update({self.param_pmc_journal_type: ' '.join(item[self.param_pmc_journal_type])})
            item.update({self.param_pmc_publication_type: _publication_type})
            item.update({self.param_pmc_title: ' '.join(item[self.param_pmc_title])})
            item.update({self.param_pmc_abstract: ' '.join(item[self.param_pmc_abstract]).replace('\"', '')})
            item.update({self.param_pmc_author: '; '.join(item[self.param_pmc_author])})

            _language_raw = item[self.param_pmc_language]
            _language = []
            for code in _language_raw:
                _language.append(self.get_language(code))
            item.update({self.param_pmc_language: ' '.join(_language)})

            _date = ' '.join(item[self.param_pmc_date])
            if _date:
                _date = _date[0:4]
            item.update({self.param_pmc_date: _date})

            _doi_raw = ' '.join(item[self.param_pmc_doi])
            _doi_raw = _doi_raw.split('doi:')
            _doi = ''
            if len(_doi_raw) > 1:
                _doi = self.remove_endpoint(_doi_raw[1])
            item.update({self.param_pmc_doi: _doi})

        # Temporary file .csv
        fw_tmp = tempfile.NamedTemporaryFile(mode = 'w+t',
                                             encoding = 'utf-8',
                                             prefix = 'medline_output_',
                                             suffix = '.csv',
                                             delete = False)

        fw_tmp.write('"%s","%s","%s","%s","%s","%s","%s","%s","%s","%s"\n' % ('PMID',
                                                                              self.pmc_col_title,
                                                                              self.pmc_col_authors,
                                                                              self.pmc_col_year,
                                                                              'PMCID',
                                                                              self.pmc_col_doi,
                                                                              self.pmc_col_language,
                                                                              self.pmc_col_document_type,
                                                                              'Journal Type',
                                                                              self.pmc_col_abstract))
        for _, detail in medline_data.items():
            fw_tmp.write('"%s","%s","%s","%s","%s","%s","%s","%s","%s","%s"\n' % (detail[self.param_pmc_pmid],
                                                                                  detail[self.param_pmc_title],
                                                                                  detail[self.param_pmc_author],
                                                                                  detail[self.param_pmc_date],
                                                                                  detail[self.param_pmc],
                                                                                  detail[self.param_pmc_doi],
                                                                                  detail[self.param_pmc_language],
                                                                                  detail[self.param_pmc_publication_type],
                                                                                  detail[self.param_pmc_journal_type],
                                                                                  detail[self.param_pmc_abstract]))
        fw_tmp.seek(0)
        # fw_tmp.close()

        return fw_tmp.name

    def read_embase_file(self, file):
        # Check format
        # Temporary file .csv
        fw_tmp = tempfile.NamedTemporaryFile(mode = 'w+t',
                                             encoding = 'utf-8',
                                             prefix = 'embase_output_',
                                             suffix = '.csv',
                                             delete = False)

        flag_index = 0
        with open(file, 'r', encoding = 'utf-8') as fr:
            for index, line in enumerate(fr):
                if 'SEARCH QUERY' in line:
                    flag_index = 3

                if index >= flag_index:
                    fw_tmp.write(line)
        fr.close()
        fw_tmp.seek(0)
        # fw_tmp.close()

        return fw_tmp.name

    def read_dimensions_file(self, file):
        # Check format
        # Temporary file .csv
        fw_tmp = tempfile.NamedTemporaryFile(mode = 'w+t',
                                             encoding = 'utf-8',
                                             prefix = 'dimensions_output_',
                                             suffix = '.csv',
                                             delete = False)

        with open(file, 'r', encoding = 'utf-8') as fr:
            for line in fr:
                flag_save = True
                if 'About the data: Exported on' in line and 'Criteria:' in line:
                    flag_save = False

                if flag_save:
                    fw_tmp.write(line)
        fr.close()
        fw_tmp.seek(0)
        # fw_tmp.close()

        return fw_tmp.name

    def read_sciencedirect_file(self, file):
        sciencedirect_data = []
        current_record = {}
        authors = []
        keywords = []

        with open(file, 'r', encoding = 'utf-8') as fr:
            for line in fr:
                line = line.strip()
                if not line:
                    continue

                if line.startswith('TY  -'):
                    if current_record:
                        current_record[self.param_sciencedirect_authors] = '; '.join(authors)
                        current_record[self.param_sciencedirect_keywords] = '; '.join(keywords)
                        sciencedirect_data.append(current_record)

                    current_record = {'TY': line.split('- ')[1].strip()}
                    authors = []
                    keywords = []
                elif line.startswith('AU  -'):
                    arr_line = line.split('- ')
                    if len(arr_line) > 1:
                        author = arr_line[1].strip()
                        authors.append(author)
                elif line.startswith('KW  -'):
                    arr_line = line.split('- ')
                    if len(arr_line) > 1:
                        keyword = arr_line[1].strip()
                        keywords.append(keyword)
                else:
                    key_value = line.split('- ', 1)
                    if len(key_value) == 2:
                        key, value = key_value
                        key = key.strip()
                        value = value.strip()
                        current_record[key] = value

        if current_record:
            current_record[self.param_sciencedirect_authors] = '; '.join(authors)
            current_record[self.param_sciencedirect_keywords] = '; '.join(keywords)
            sciencedirect_data.append(current_record)

        fw_tmp = tempfile.NamedTemporaryFile(mode = 'w+t',
                                             encoding = 'utf-8',
                                             prefix = 'sciencedirect_output_',
                                             suffix = '.csv',
                                             delete = False)

        headers = [self.param_sciencedirect_title,
                   self.param_sciencedirect_authors,
                   self.param_sciencedirect_journal,
                   self.param_sciencedirect_publication_year,
                   self.param_sciencedirect_volume,
                   self.param_sciencedirect_first_page,
                   self.param_sciencedirect_last_page,
                   self.param_sciencedirect_publication_date,
                   self.param_sciencedirect_issn,
                   self.param_sciencedirect_abstract,
                   self.param_sciencedirect_keywords,
                   self.param_sciencedirect_doi,
                   self.param_sciencedirect_sciencedirect_link]

        fw_tmp.write('"' + '"|"'.join(headers) + '"\n')

        for record in sciencedirect_data:
            row = [record.get(self.param_sciencedirect_title, '').replace('“', '').replace('”', ''),
                   record.get(self.param_sciencedirect_authors, ''),
                   record.get(self.param_sciencedirect_journal, ''),
                   record.get(self.param_sciencedirect_publication_year, ''),
                   record.get(self.param_sciencedirect_volume, ''),
                   record.get(self.param_sciencedirect_first_page, ''),
                   record.get(self.param_sciencedirect_last_page, ''),
                   record.get(self.param_sciencedirect_publication_date, ''),
                   record.get(self.param_sciencedirect_issn, ''),
                   record.get(self.param_sciencedirect_abstract, '').replace('"', ''),
                   record.get(self.param_sciencedirect_keywords, ''),
                   record.get(self.param_sciencedirect_doi, '').replace('https://doi.org/', ''),
                   record.get(self.param_sciencedirect_sciencedirect_link, '')]

            fw_tmp.write('"' + '"|"'.join(row) + '"\n')

        fw_tmp.seek(0)

        return fw_tmp.name

def main():
    try:
        start = ofi.start_time()
        menu()

        ofi.LOG_FILE = os.path.join(ofi.OUTPUT_PATH, ofi.LOG_NAME)
        ofi.XLS_FILE = os.path.join(ofi.OUTPUT_PATH, ofi.XLS_FILE.replace('<type>', ofi.TYPE_FILE))
        ofi.show_print("#############################################################################", [ofi.LOG_FILE], font = ofi.BIGREEN)
        ofi.show_print("############################### Format Input ################################", [ofi.LOG_FILE], font = ofi.BIGREEN)
        ofi.show_print("#############################################################################", [ofi.LOG_FILE], font = ofi.BIGREEN)

        # Read input file
        input_information = {}
        if ofi.TYPE_FILE == ofi.TYPE_TXT:
            ofi.show_print("Reading the .txt file", [ofi.LOG_FILE], font = ofi.GREEN)
            input_information = ofi.read_txt_file()
        elif ofi.TYPE_FILE == ofi.TYPE_SCOPUS:
            ofi.show_print("Reading the .csv file from Scopus", [ofi.LOG_FILE], font = ofi.GREEN)
            input_information = ofi.read_csv_file()
        elif ofi.TYPE_FILE in [ofi.TYPE_WOS, ofi.TYPE_SCIELO]:
            database = 'Web of Science'
            if ofi.TYPE_FILE == ofi.TYPE_SCIELO:
                database = 'SciELO'
            ofi.show_print("Reading the .csv file from %s" % database, [ofi.LOG_FILE], font = ofi.GREEN)
            input_information = ofi.read_csv_file()
        elif ofi.TYPE_FILE == ofi.TYPE_PUBMED:
            ofi.show_print("Reading the .csv file from PubMed", [ofi.LOG_FILE], font = ofi.GREEN)
            input_information = ofi.read_csv_file()
        elif ofi.TYPE_FILE == ofi.TYPE_PUBMED_CENTRAL:
            ofi.show_print("Reading the .txt file from PubMed Central", [ofi.LOG_FILE], font = ofi.GREEN)
            input_information = ofi.read_csv_file()
        elif ofi.TYPE_FILE == ofi.TYPE_DIMENSIONS:
            ofi.show_print("Reading the .csv file from Dimensions", [ofi.LOG_FILE], font = ofi.GREEN)
            input_information = ofi.read_csv_file()
        elif ofi.TYPE_FILE == ofi.TYPE_GOOGLE_SCHOLAR:
            ofi.show_print("Reading the .csv file from Publish or Perish (Google Scholar option)", [ofi.LOG_FILE], font = ofi.GREEN)
            input_information = ofi.read_csv_file()
        elif ofi.TYPE_FILE == ofi.TYPE_COCHRANE:
            ofi.show_print("Reading the .csv file from Cochrane", [ofi.LOG_FILE], font = ofi.GREEN)
            input_information = ofi.read_csv_file()
        elif ofi.TYPE_FILE == ofi.TYPE_EMBASE:
            ofi.show_print("Reading the .csv file from Embase", [ofi.LOG_FILE], font = ofi.GREEN)
            input_information = ofi.read_csv_file()
        elif ofi.TYPE_FILE == ofi.TYPE_SCIENCEDIRECT:
            ofi.show_print("Reading the .ris file from ScienceDirect", [ofi.LOG_FILE], font = ofi.GREEN)
            input_information = ofi.read_csv_file()
        elif ofi.TYPE_FILE == ofi.TYPE_IEEE:
            ofi.show_print("Reading the .csv file from IEEE", [ofi.LOG_FILE], font = ofi.GREEN)
            input_information = ofi.read_csv_file()
        elif ofi.TYPE_FILE == ofi.TYPE_BVS:
            ofi.show_print("Reading the .csv file from BVS", [ofi.LOG_FILE], font = ofi.GREEN)
            input_information = ofi.read_csv_file()
        elif ofi.TYPE_FILE == ofi.TYPE_CAB:
            ofi.show_print("Reading the .csv file from CAB", [ofi.LOG_FILE], font = ofi.GREEN)
            input_information = ofi.read_csv_file()
        else:
            ofi.show_print("%s: error: Option not found '%s'" % (os.path.basename(__file__), ofi.TYPE_FILE), showdate = False, font = ofi.YELLOW)
            exit()
        # pprint(input_information)

        ofi.show_print("Input file: %s" % ofi.INPUT_FILE, [ofi.LOG_FILE])
        ofi.show_print("", [ofi.LOG_FILE])

        ofi.save_summary_xls(input_information)
        ofi.show_print("Output file: %s" % ofi.XLS_FILE, [ofi.LOG_FILE], font = ofi.GREEN)
        ofi.show_print("  Unique documents: %s" % len(input_information[ofi.XLS_SHEET_UNIQUE]), [ofi.LOG_FILE])
        ofi.show_print("  Duplicate documents: %s" % len(input_information[ofi.XLS_SHEET_DUPLICATES]), [ofi.LOG_FILE])
        if ofi.TYPE_FILE != ofi.TYPE_TXT:
            ofi.show_print("  Documents without DOI: %s" % len(input_information[ofi.XLS_SHEET_WITHOUT_DOI]), [ofi.LOG_FILE])

        ofi.show_print("", [ofi.LOG_FILE])
        ofi.show_print(ofi.finish_time(start, "Elapsed time"), [ofi.LOG_FILE])
        ofi.show_print("Done!", [ofi.LOG_FILE])
    except Exception as e:
        ofi.show_print("\n%s" % traceback.format_exc(), [ofi.LOG_FILE], font = ofi.RED)
        ofi.show_print(ofi.finish_time(start, "Elapsed time"), [ofi.LOG_FILE])
        ofi.show_print("Done!", [ofi.LOG_FILE])

if __name__ == '__main__':
    ofi = FormatInput()
    main()
