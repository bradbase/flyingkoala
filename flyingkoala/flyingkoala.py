
import logging

import xlwings as xw
from xlcalculator import ModelCompiler
from xlcalculator import Model
from xlcalculator import Evaluator
import numpy as np
import pandas as pd

logging.basicConfig(level=logging.INFO)

class FlyingKoala():

    excel_file_name = None
    auto_load_value = False
    koala_models = {}
    excel_model = None
    ignore_sheets_default = "FlyingKoala.conf\\_FlyingKoala.conf\\xlwings.conf\\_xlwings.conf"
    ignore_sheets = None


    def __init__(self, file_name, ignore_sheets=[], load_koala=False):

        workbook = None

        print("len(xw.apps)", len(xw.apps))

        if len(xw.apps) != 0:
            books = [book.name for book in xw.books]
            print("books", books, file_name)
            for book in books:
                if book in file_name:
                    workbook = xw.books[book]

        else:
            xw.books.open(file_name)
            workbook = xw.books[file_name]

        print("workbook", workbook)

        if 'FlyingKoala.conf' in [ sheet.name for sheet in workbook.sheets ]:
            self.auto_load_value = workbook.sheets['FlyingKoala.conf'].range('B3').value
        else:
            self.auto_load_value = False
            logging.error("Would be great to see a worksheet FlyingKoala.conf")

        self.excel_file_name = workbook.fullname
        self.ignore_sheets = "{}\\{}".format(workbook.sheets['FlyingKoala.conf'].range('B2').value, self.ignore_sheets_default)

        if self.auto_load_value == True or load_koala == True:
            self.reload_koala(self.excel_file_name, ignore_sheets=self.ignore_sheets)


    def generate_model_graph(self, model, refresh = False):
        """The function that extracts a graph of a given model from the Spreadsheet"""

        if isinstance(model, str):
            if refresh == False and model in self.koala_models.keys():
                return 'Model %s is already cached, set refresh True if you want it to refresh it' % model.name.name
        else:
            if refresh == False and model.name.name in self.koala_models.keys():
                return 'Model %s is already cached, set refresh True if you want it to refresh it' % model.name.name

        # logging.debug(parser.prettyprint())
        inputs = [model.name.name]
        inputs.extend(self.excel_model.formulae[model.name.name].terms)
        self.koala_models[str(model.name.name)] = ModelCompiler.extract(self.excel_model, focus=inputs)

        logging.info("Successfully loaded model {}".format(model.name))
        return 'Cached Model %s' % model.name


    def reload_koala(self, file_name, ignore_sheets= None, bootstrap_equations= None):
        """Loads the Excel workbook into a Python compatible object"""

        if ignore_sheets is not None:
            ignore_sheets = ignore_sheets.split('\\')
        else:
            ignore_sheets = self.ignore_sheets.split('\\')

        logging.info("Loading workbook")

        if file_name == '':
            logging.debug("file_name is not set in Excel Ribbon using {}".format(self.excel_file_name))
            file_name = self.excel_file_name

        self.excel_compiler = ModelCompiler()
        self.excel_model = self.excel_compiler.read_and_parse_archive(file_name, ignore_sheets=self.ignore_sheets)

        logging.info("Workbook '{}' has been loaded.".format(file_name))
        logging.info("Ignored worksheets {}".format(ignore_sheets))


    def reset_koala_model_cache(self):
        global koala_models

        self.koala_models = {}


    def get_model_cache_count(self):
        return len(self.koala_models.keys())


    @staticmethod
    def get_named_range_count():
        wb = xw.books.active
        return len(wb.names)


    def get_cached_koala_model_names(self):

        names_of_cached_models = ""
        for model_name in self.koala_models.keys():
            names_of_cached_models += "\r\n%s" % (model_name)

        return names_of_cached_models


    @staticmethod
    def get_named_ranges():

        returnable = ""
        wb = xw.books.active
        for name in wb.names:
            returnable += "\r\n%s" % (name.name)

        return returnable


    def is_koala_model_cached(self, model_name):

        return model_name in self.koala_models.keys()


    def load_model(self, model_name):
        """Preparing model name from either a string or an xlwings Range and load it into cache."""

        if self.excel_model is None:
            self.reload_koala(self.excel_file_name, ignore_sheets=self.ignore_sheets)

        # figure out if we have a named range or a text name of the model
        extracted_model_name = None
        if model_name is not None:
            if model_name.name is None:
                extracted_model_name = model_name.value
            else:
                extracted_model_name = model_name.name.name
        else:
            raise Exception('The range you tried to use does not exist in the workbook, if named range exists check spelling.')

        # ensure that model is cached
        if extracted_model_name not in self.koala_models.keys():
            model = None
            wb = xw.books.active
            for name in wb.names:
                if extracted_model_name == name.name:
                    model = xw.Range(extracted_model_name)
                    self.generate_model_graph(model)

            if model is None:
                raise Exception('Model "%s" has not been loaded into cache, if named range exists check spelling.' % extracted_model_name)

        return extracted_model_name


    def unload_koala_model_from_cache(self, model_name):
        global koala_models

        del(self.koala_models[model_name])


    def evaluate_koala_model_row(self, model_name, input_data, model, no_calc_when_zero=[]):
        """The function which sets the input values in the model and evaluates the Excel equation using koala"""

        model = self.koala_models[str(model_name)]
        for key in input_data.keys():
            if key in no_calc_when_zero and input_data[key] == 0:
                return

            model.set_cell_value(key, input_data[key])

        evaluator = Evaluator(model)

        return evaluator.evaluate(model_name)


    def evaluate_koala_model(self, model_name, terms, no_calc_when_zero=[]):
        """The function that sets up the evaluation of the koala equation"""

        if model_name not in self.koala_models.keys():
            return 'Model %s has not been loaded into cache.' % model_name

        def eval(row, model):
            return self.evaluate_koala_model_row(model_name, row.to_dict(), model, no_calc_when_zero)

        current_model = self.koala_models[model_name]

        results = terms.apply(eval, axis= 1, model= current_model)
        return results



workbook = xw.books.active
excel_file_name = workbook.fullname
workbook.sheets['FlyingKoala.conf'].range('B2').value
if 'FlyingKoala.conf' in [ sheet.name for sheet in workbook.sheets ]:
    auto_load_value = workbook.sheets['FlyingKoala.conf'].range('B3').value
else:
    auto_load_value = False
    logging.error("Would be great to see a worksheet FlyingKoala.conf")

ignore_sheets_default = "FlyingKoala.conf\\_FlyingKoala.conf\\xlwings.conf\\_xlwings.conf"
ignore_sheets = "{}\\{}".format(workbook.sheets['FlyingKoala.conf'].range('B2').value, ignore_sheets_default)

my_fk = FlyingKoala(excel_file_name, ignore_sheets, auto_load_value)


@xw.sub
def generate_model_graph(model, refresh = False):
    """The function that extracts a graph of a given model from the Spreadsheet"""
    return my_fk.generate_model_graph(model, refresh=refresh)


@xw.sub
def reload_koala(file_name, ignore_sheets= None, bootstrap_equations= None):
    """Loads the Excel workbook into a Python compatible object"""
    my_fk.reload_koala(file_name, ignore_sheets=ignore_sheets, bootstrap_equations=bootstrap_equations)


@xw.func
def reset_koala_model_cache():
    my_fk.reset_koala_model_cache()


@xw.func
def get_model_cache_count():
    return my_fk.get_model_cache_count()


@xw.func
def get_named_range_count():
    return my_fk.get_named_range_count()


@xw.func
def get_cached_koala_model_names():
    return my_fk.get_cached_koala_model_names()


@xw.func
def get_named_ranges():
    return my_fk.get_named_ranges()


@xw.func
@xw.arg('model_name', doc='Name, as a string, of the model which might be cached.')
def is_koala_model_cached(model_name):
    return my_fk.is_koala_model_cached(model_name)


def load_model(model_name):
    """Preparing model name from either a string or an xlwings Range and load it into cache."""
    return my_fk.load_model(model_name)


@xw.func
@xw.arg('model_name', doc='Name, as a string, of the model which might be cached.')
def unload_koala_model_from_cache(model_name):
    my_fk.unload_koala_model_from_cache(model_name)


@xw.sub
def evaluate_koala_model(model_name, terms, no_calc_when_zero=[]):
    """The function that sets up the evaluation of the koala equation"""
    return my_fk.evaluate_koala_model(model_name, terms, no_calc_when_zero=no_calc_when_zero)
