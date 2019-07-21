
import xlwings as xw
import numpy as np
import pandas as pd

from flyingkoala import flyingkoala as fk

@xw.func
@xw.arg('model', xw.Range, doc='Name, as a string, of the model which will be evaluated. The Excel cell name / named range')
@xw.arg('T_min', np.array, doc='Daily minimum temperature')
@xw.arg('T_max', np.array, doc='Daily maximum temperature')
@xw.ret(index=False, header=False)
def DegreeDay(model, T_min, T_max):
    """Function to assemble a dataframe for calculating Degree Day"""

    if not fk.isKoalaModelCached(model.name.name):
        fk.generateModelGraph(model)

    inputs_for_DegreeDay = pd.DataFrame({'T_min': np.array([T_min]), 'T_max': np.array([T_max])})
    return fk.EvaluateKoalaModel(model.name.name, inputs_for_DegreeDay)

@xw.func
@xw.arg('model', xw.Range, doc='Name, as a string, of the model which will be evaluated. The Excel cell name / named range.')
@xw.arg('T_min', np.array, doc='Daily minimum temperature')
@xw.arg('T_max', np.array, doc='Daily maximum temperature')
@xw.ret(index=False, header=False, expand='down')
def DegreeDayDynamicArray(model, T_min, T_max):
    """Function to assemble a dataframe for calculating Degree Day using dynamic arrays"""

    if not fk.isKoalaModelCached(model.name.name):
        fk.generateModelGraph(model)

    inputs_for_DegreeDay = pd.DataFrame({'T_min': T_min, 'T_max': T_max})
    return fk.EvaluateKoalaModel(model.name.name, inputs_for_DegreeDay)
