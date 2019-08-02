import xlwings as xw
import numpy as np
import pandas as pd

from flyingkoala import flyingkoala as fk

@xw.func
@xw.arg('model', xw.Range, doc='Named Range of the model which will be evaluated. The Excel cell name / named range')
@xw.arg('T_min', np.array, doc='Daily minimum temperature')
@xw.arg('T_max', np.array, doc='Daily maximum temperature')
@xw.ret(index=False, header=False)
def DegreeDay(model, T_min, T_max):
    """Function to assemble a dataframe for calculating Degree Day"""

    if not fk.is_koala_model_cached(model.name.name):
        fk.generate_model_graph(model)

    inputs_for_DegreeDay = pd.DataFrame({'T_min': np.array([T_min]), 'T_max': np.array([T_max])})
    return fk.evaluate_koala_model(model.name.name, inputs_for_DegreeDay)


@xw.func
@xw.arg('model', xw.Range, doc='Named Range of the model which will be evaluated. The Excel cell name / named range.')
@xw.arg('T_min', np.array, doc='Daily minimum temperature')
@xw.arg('T_max', np.array, doc='Daily maximum temperature')
@xw.ret(index=False, header=False, expand='down')
def DegreeDayDynamicArray(model, T_min, T_max):
    """Function to assemble a dataframe for calculating Degree Day using dynamic arrays"""

    if not fk.is_koala_model_cached(model.name.name):
        fk.generate_model_graph(model)

    inputs_for_DegreeDay = pd.DataFrame({'T_min': T_min, 'T_max': T_max})
    return fk.evaluate_koala_model(model.name.name, inputs_for_DegreeDay)


@xw.func
@xw.arg('model_name', doc='Name, as a string, of the model which will be evaluated. The Excel cell name / named range')
@xw.arg('T_min', np.array, doc='Daily minimum temperature')
@xw.arg('T_max', np.array, doc='Daily maximum temperature')
@xw.ret(index=False, header=False)
def DegreeDayModelByName(model_name, T_min, T_max):
    """Function to assemble a dataframe for calculating Degree Day"""

    if model_name not in fk.koala_models.keys():
        wb = xw.books.active
        for name in wb.names:
            if model_name == name.name:
                model = xw.Range(model_name)
                fk.generate_model_graph(model)

        if model is None:
            return 'Model %s has not been loaded into cache, if named range exists check spelling.' % model_name

    inputs_for_DegreeDay = pd.DataFrame({'T_min': np.array([T_min]), 'T_max': np.array([T_max])})
    return fk.evaluate_koala_model(model_name, inputs_for_DegreeDay)


@xw.func
@xw.arg('model_name', doc='Name, as a string, of the model which will be evaluated. The Excel cell name / named range')
@xw.arg('T_min', np.array, doc='Daily minimum temperature')
@xw.arg('T_max', np.array, doc='Daily maximum temperature')
@xw.ret(index=False, header=False, expand='down')
def DegreeDayDynamicArrayByName(model_name, T_min, T_max):
    """Function to assemble a dataframe for calculating Degree Day using dynamic arrays"""

    if model_name not in fk.koala_models.keys():
        wb = xw.books.active
        for name in wb.names:
            if model_name == name.name:
                model = xw.Range(model_name)
                fk.generate_model_graph(model)

        if model is None:
            return 'Model %s has not been loaded into cache, if named range exists check spelling.' % model_name

    inputs_for_DegreeDay = pd.DataFrame({'T_min': T_min, 'T_max': T_max})
    return fk.evaluate_koala_model(model_name, inputs_for_DegreeDay)
