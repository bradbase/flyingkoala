
from datetime import timedelta

import xlwings as xw
import numpy as np
import pandas as pd

from flyingkoala import flyingkoala as fk

@xw.func
@xw.arg('times', np.array, doc='This is the range of times')
@xw.arg('inputs', np.array, doc='This is the value you want to average from.')
@xw.arg('window', doc='The number of elements which will be averaged.')
@xw.ret(index=False, header=False, expand='down')
def TIMESERIESWINDOWAVERAGE(times, inputs, window=5):
    """Performs a look-ahead moving average of size window on a time series for values in inputs"""
    def include_average(time, value, window_size):
        ts = pd.to_datetime(str(time))
        mymod = np.mod(int(ts.strftime('%M')), window_size)
        if mymod == np.int32(0):
            return value

    window_size = int(window)

    timeseries = pd.DataFrame({'time':times, 'value':inputs})
    upside_down = timeseries.iloc[::-1]
    upside_down['average'] = upside_down['value'].rolling(window_size).mean()
    timeseries = upside_down.iloc[::-1]

    timeseries['returnable'] = np.vectorize(include_average)(timeseries['time'], timeseries['average'], window_size)

    return timeseries['returnable']


@xw.func
@xw.arg('times', np.array, doc='This is the range of times')
@xw.arg('inputs', np.array, doc='This is the value you want to average from.')
@xw.arg('window', doc='The elements will be kept on this index.')
@xw.ret(expand='down')
def KEEPRECORDS(times, inputs, window=5):
    """Keeps records at an offset determined by window"""

    returnable = []
    ascending = True
    timeseries = pd.DataFrame({'time':times, 'value':inputs})

    delta = timedelta(minutes = window)

    if timeseries.index.is_monotonic_increasing:
        ascending = False

    goal_time = None
    for index, row in timeseries.iterrows():
        if goal_time == None:
            returnable.append([row['value']])
            if ascending:
                goal_time = row['time'] + delta
            else:
                goal_time = row['time'] - delta
        else:

            if goal_time - row['time'] == timedelta(milliseconds = 0):
                returnable.append([row['value']])
                if ascending:
                    goal_time = goal_time + delta
                else:
                    goal_time = goal_time - delta
            else:
                returnable.append([None])

    return returnable


@xw.func
@xw.arg('times', np.array, doc='This is the range of times')
@xw.arg('inputs', np.array, doc='This is the value you want to average from.')
@xw.arg('window', doc='The number of elements which will be averaged.')
@xw.arg('operation', doc='The operation by which the resample will occur.')
@xw.ret(index=False, header=False, expand='down')
def RESAMPLEMINS(times, inputs, window=5, operation='mean'):
    """Performs a look-ahead re-sample of size window with stated operation on a time series for values in inputs and does not return the time index"""
    window_size = int(window)

    timeseries = pd.DataFrame({'time':times, 'value':inputs})
    timeseries.set_index('time', inplace=True)
    upside_down = timeseries.iloc[::-1]
    if operation == 'mean':
        thing = upside_down.resample('{0}Min'.format(int(window))).mean()
    elif operation == 'sum':
        thing = upside_down.resample('{0}Min'.format(int(window))).sum()

    return thing['value']


@xw.func
@xw.arg('times', np.array, doc='This is the range of times')
@xw.arg('inputs', np.array, doc='This is the value you want to average from.')
@xw.arg('window', doc='The number of elements which will be averaged.')
@xw.arg('operation', doc='The operation by which the resample will occur.')
@xw.ret(index=True, header=False, expand='down')
def RESAMPLEMINSWITHINDEX(times, inputs, window=5, operation='mean'):
    """Performs a look-ahead re-sample average of size window on a time series for values in inputs and returns with the time index"""
    window_size = int(window)

    timeseries = pd.DataFrame({'time':times, 'value':inputs})
    timeseries.set_index('time', inplace=True)
    # upside_down = timeseries.iloc[::-1]
    thing = upside_down.resample('{0}Min'.format(int(window))).mean()

    return thing['value']


@xw.func
@xw.arg('keys', np.array, doc='The reference time.')
@xw.arg('below', doc='Integer number of hours leading up to the key time.')
@xw.arg('above', doc='Integer number of hours beyond the key time.')
@xw.arg('periods', np.array, doc='The times which will be determined within the period')
@xw.ret(index=False, header=False, expand='down')
def TIMEISBETWEEN(keys, below, above, periods):
    """Decides if a time is between certain range of a given time"""

    def include_period(key, below_delta, above_delta, period):
        below_date = pd.to_datetime(str(key)) - below_delta
        above_date = pd.to_datetime(str(key)) + above_delta
        if pd.to_datetime(str(period)) >= below_date and pd.to_datetime(str(period)) < above_date:
            return True

    below_delta = timedelta(hours=below)
    above_delta = timedelta(hours=above)

    time_between = pd.DataFrame({'keys': keys, 'periods': periods})

    time_between['returnable'] = np.vectorize(include_period)(time_between['keys'], below_delta, above_delta, time_between['periods'])

    return time_between['returnable']
