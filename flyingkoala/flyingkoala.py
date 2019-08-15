
import xlwings as xw
from koala.ExcelCompiler import ExcelCompiler
from koala.tokenizer import ExcelParser
from koala.Spreadsheet import Spreadsheet
import numpy as np
import pandas as pd

workbook = xw.books.active
auto_load = workbook.sheets['FlyingKoala.conf'].range('B3')

excel_file_name = None
excel_compiler = None
ignore_sheets = []
koala_models = {}

def parse_model(model):
    """Parses an Excel formula into tokens and returns the operand ranges.

    :param model: A text representation of an Excel formula.
    """

    parser = ExcelParser()
    tokens = parser.parse(model.formula)
    print(parser.prettyprint())
    return parser.getOperandRanges()

@xw.sub
def generate_model_graph(model, refresh = False):
    """The function that extracts a graph of a given model from the Spreadsheet.

    :param model: An xlwings Range object with an equation in it.
    :param refresh: A flag to indicate whether an existing model needs to be re-loaded"""
    global koala_models
    global excel_compiler

    if isinstance(model, str):
        if refresh == False and model in koala_models.keys():
            return 'Model %s is already cached, set refresh True if you want it to refresh it' % model.name.name
    else:
        if refresh == False and model.name.name in koala_models.keys():
            return 'Model %s is already cached, set refresh True if you want it to refresh it' % model.name.name

    inputs = parse_model(model)
    # koala_models[str(model.name.name)] = excel_compiler.gen_graph(inputs= inputs, outputs= [model.name.name])
    koala_models[str(model.name.name)] = excel_compiler.gen_graph()

    print("Successfully loaded model %s" % model.name)
    return 'Cached Model %s' % model.name


@xw.sub
def reload_koala(file_name, ignore_sheets= None, bootstrap_equations= None):
    """Loads the Excel workbook into a koala Spreadsheet object"""
    global excel_compiler
    print("Loading workbook")
    excel_compiler = Spreadsheet.from_file_name(file_name, ignore_sheets=ignore_sheets)
    excel_compiler.clean_pointer()
    print("Workbook '%s' has been loaded." % file_name)
    print("Ignored worksheets %s" % ignore_sheets)


if auto_load.value == True:
    excel_file_name = workbook.fullname
    reload_koala(excel_file_name, ignore_sheets=ignore_sheets)

@xw.func
def reset_koala_model_cache():
    global koala_models
    koala_models = {}

@xw.func
def get_model_cache_count():
    global koala_models

    return len(koala_models.keys())

@xw.func
def get_named_range_count():

    wb = xw.books.active

    return len(wb.names)

@xw.func
def get_cached_koala_model_names():
    global koala_models

    names_of_cached_models = ""
    for model_name in koala_models.keys():
        names_of_cached_models += "\r\n%s" % (model_name)

    return names_of_cached_models

@xw.func
def get_named_ranges():

    returnable = ""
    wb = xw.books.active

    for name in wb.names:
        returnable += "\r\n%s" % (name.name)

    return returnable

@xw.func
@xw.arg('model_name', doc='Name, as a string, of the model which might be cached.')
def is_koala_model_cached(model_name):
    global koala_models

    return model_name in koala_models.keys()

def load_model(model_name):
    """Preparing model name from either a string or an xlwings Range and load it into cache."""
    global koala_models

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
    if extracted_model_name not in koala_models.keys():
        model = None
        wb = xw.books.active
        for name in wb.names:
            if extracted_model_name == name.name:
                model = xw.Range(extracted_model_name)
                generate_model_graph(model)

        if model is None:
            raise Exception('Model "%s" has not been loaded into cache, if named range exists check spelling.' % extracted_model_name)

    return extracted_model_name

@xw.func
@xw.arg('model_name', doc='Name, as a string, of the model which might be cached.')
def unload_koala_model_from_cache(model_name):
    global koala_models
    del(koala_models[model_name])

def evaluate_koala_model_row(model_name, input_data, model, no_calc_when_zero=[]):
    """The function which sets the input values in the model and evaluates the Excel equation using koala"""
    global koala_models
    model = koala_models[str(model_name)]
    for key in input_data.keys():
        if key in no_calc_when_zero and input_data[key] == 0:
            return

        print("key", key, "data", input_data[key])
        model.set_value(key, input_data[key])

    return model.evaluate(model_name)

@xw.sub
def evaluate_koala_model(model_name, terms, no_calc_when_zero=[]):
    """The function that sets up the evaluation of the koala equation"""
    global koala_models

    if model_name not in koala_models.keys():
        return 'Model %s has not been loaded into cache.' % model_name

    def eval(row, model):
        return evaluate_koala_model_row(model_name, row.to_dict(), model, no_calc_when_zero)

    current_model = koala_models[model_name]

    results = terms.apply(eval, axis= 1, model= current_model)
    return results
