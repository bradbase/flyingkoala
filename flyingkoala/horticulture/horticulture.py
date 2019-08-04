import xlwings as xw
import numpy as np
import pandas as pd

from flyingkoala import flyingkoala as fk

@xw.func
@xw.arg('model_name', doc='Name, as a string, of the model which will be evaluated. The Excel cell name / named range')
@xw.arg('T_min', np.array, doc='Daily minimum temperature')
@xw.arg('T_max', np.array, doc='Daily maximum temperature')
@xw.ret(index=False, header=False)
def DegreeDayByName(model_name, T_min, T_max):
    """Function to assemble a dataframe for calculating Degree Day.

    :param model: Text name (string) of the model which will be evaluated. The Excel cell name / named range
    :param T_min: Daily minimum temperature
    :param T_max: Daily maximum temperature
    """

    if model_name not in fk.koala_models.keys():
        model = None
        wb = xw.books.active
        for name in wb.names:
            if model_name == name.name:
                model = xw.Range(model_name)
                fk.generate_model_graph(model)

        if model is None:
            return 'Model "%s" has not been loaded into cache, if named range exists check spelling.' % model_name

    inputs_for_DegreeDay = pd.DataFrame({'T_min': np.array([T_min]), 'T_max': np.array([T_max])})
    return fk.evaluate_koala_model(model_name, inputs_for_DegreeDay)

@xw.func
@xw.arg('model_name', doc='Name, as a string, of the model which will be evaluated. The Excel cell name / named range')
@xw.arg('T_min', np.array, doc='Daily minimum temperature')
@xw.arg('T_max', np.array, doc='Daily maximum temperature')
@xw.ret(index=False, header=False, expand='down')
def DegreeDayDynamicArrayByName(model_name, T_min, T_max):
    """Function to assemble a dataframe for calculating Degree Day using dynamic arrays.

    :param model: Text name (string) of the model which will be evaluated. The Excel cell name / named range
    :param T_min: Daily minimum temperature
    :param T_max: Daily maximum temperature
    """


    if model_name not in fk.koala_models.keys():
        model = None
        wb = xw.books.active
        for name in wb.names:
            if model_name == name.name:
                model = xw.Range(model_name)
                fk.generate_model_graph(model)

        if model is None:
            return 'Model "%s" has not been loaded into cache, if named range exists check spelling.' % model_name

    inputs_for_DegreeDay = pd.DataFrame({'T_min': T_min, 'T_max': T_max})
    return fk.evaluate_koala_model(model_name, inputs_for_DegreeDay)

@xw.func
@xw.arg('model_name', xw.Range, doc='Name, as a string or named range, of the model which will be evaluated.')
@xw.arg('T_min', np.array, doc='Daily minimum temperature')
@xw.arg('T_max', np.array, doc='Daily maximum temperature')
@xw.ret(index=False, header=False)
def DegreeDay(model_name, T_min, T_max):
    """Function to assemble a dataframe for calculating Degree Day. Takes either a name of a model or the named range itself.

    :param model: Named Range of the model which will be evaluated. The Excel cell name / named range
    :param T_min: Daily minimum temperature
    :param T_max: Daily maximum temperature
    """
    try:
        prepared_model_name = fk.load_model(model_name)

        # format the equation terms, evalueate the equation and return the result.
        inputs_for_DegreeDay = pd.DataFrame({'T_min': np.array([T_min]), 'T_max': np.array([T_max])})
        return fk.evaluate_koala_model(prepared_model_name, inputs_for_DegreeDay)

    except Exception as err:
        return str(err)


@xw.func
@xw.arg('model_name', xw.Range, doc='Name, as a string or named range, of the model which will be evaluated.')
@xw.arg('T_min', np.array, doc='Daily minimum temperature')
@xw.arg('T_max', np.array, doc='Daily maximum temperature')
@xw.ret(index=False, header=False, expand='down')
def DegreeDayDynamicArray(model_name, T_min, T_max):
    """Function to assemble a dataframe for calculating Degree Day using dynamic arrays. Takes either a name of a model or the named range itself.

    :param model: Named Range of the model which will be evaluated. The Excel cell name / named range
    :param T_min: Daily minimum temperature
    :param T_max: Daily maximum temperature
    """
    try:
        prepared_model_name = fk.load_model(model_name)

        # format the equation terms, evalueate the equation and return the result.
        inputs_for_DegreeDay = pd.DataFrame({'T_min': T_min, 'T_max': T_max})
        return fk.evaluate_koala_model(prepared_model_name, inputs_for_DegreeDay)

    except Exception as err:
            return str(err)
