
from datetime import *

import xlwings as xw
import numpy as np
import pandas as pd
import pvlib
from timezonefinder import TimezoneFinder
import pytz

import flyingkoala
from flyingkoala import *

timezone_finder = TimezoneFinder()

@xw.func
@xw.arg('latitude', doc='The latitude of the desired location')
@xw.arg('longitude', doc='The longitude of the desired location')
@xw.arg('times', np.array, doc='This is the range of times')
@xw.ret(index=False, header=False, transpose=True, expand='down')
def SOLARNOON(latitude, longitude, times):
    """Finds solar noon for each day in teh series"""
    # global timezone_finder
    # timezone_name = timezone_finder.timezone_at(lng=longitude, lat=latitude)
    # tz = timezone(timezone_name)

    float_lat = float(latitude)
    float_lon = float(longitude)

    if not isinstance(times[0], datetime):
        times = pd.to_datetime(times)

    if times[0].tzinfo is not None and times[0].tzinfo.utcoffset(times[0]) is not None:
        print("timezone aware")
    else:
        print("timezone naieve")

    solar_noon = pd.DataFrame({'time_local':times})

    day_lookup = []
    def days(time):
        date_only = pd.to_datetime(str(time)).strftime('%Y-%m-%d')
        if date_only not in day_lookup:
            day_lookup.append(date_only)
        return date_only
    solar_noon['date_local'] = np.vectorize(days)(solar_noon['time_local'])

    solar_noons = {}
    for day in day_lookup:
        this_days_series = solar_noon[ solar_noon['date_local'] == day ]
        solar_position = pvlib.solarposition.get_solarposition(this_days_series['time_local'], float_lat, float_lon, method='nrel_numpy')
        solar_noons[day] = solar_position['zenith'].idxmin()
    solar_noon['zenith_local'] = solar_noon['date_local'].map(solar_noons)

    def drop_timezone(zenith_utc):
        return zenith_utc.replace(tzinfo=None)

    solar_noon['zenith_local_sans_tz'] = solar_noon['zenith_local'].apply(drop_timezone)

    return np.vectorize(pd.Timestamp)(solar_noon['zenith_local_sans_tz'].values)


@xw.func
@xw.arg('model_name', xw.Range, doc='Name, as a string, of the model which will be assessed')
@xw.arg('clean_panel_short_circuit', np.array, doc='')
@xw.arg('clean_panel_temp', np.array, doc='')
@xw.arg('clean_Isc_0', doc='')
@xw.ret(index=False, header=False, expand='down')
def SOILING_G(model_name, clean_panel_short_circuit, clean_panel_temp, clean_Isc_0):
    """Convenience function for assemble a dataframe for calculating soiling value 'G'"""

    if not isKoalaModelCached(model_name.name.name):
        generateModelGraph(model_name)

    inputs_for_G = pd.DataFrame({'clean_panel_short_circuit': clean_panel_short_circuit, 'clean_panel_temp': clean_panel_temp, 'clean_Isc_0': clean_Isc_0})
    return EvaluateKoalaModel(model_name.name.name, inputs_for_G)

@xw.func
@xw.arg('model_name', xw.Range, doc='Name, as a string, of the model which will be assessed')
@xw.arg('G_value', np.array, doc='')
@xw.arg('dirty_panel_short_circut', np.array, doc='')
@xw.arg('dirty_panel_temp', np.array, doc='')
@xw.arg('dirty_Isc_0', doc='')
@xw.ret(index=False, header=False, expand='down')
def SOILING_SR_Isc(model_name, G_value, dirty_panel_short_circut, dirty_panel_temp, dirty_Isc_0):
    """Convenience function to assemble a dataframe for calculating soiling value 'SR_(Isc)'"""

    if not isKoalaModelCached(model_name.name.name):
        generateModelGraph(model_name)

    inputs_for_SR_Isc = pd.DataFrame({'G_value': G_value, 'dirty_panel_short_circut': dirty_panel_short_circut, 'dirty_panel_temp': dirty_panel_temp, 'dirty_Isc_0': dirty_Isc_0})
    return EvaluateKoalaModel(model_name.name.name, inputs_for_SR_Isc)

@xw.func
@xw.arg('model_name', xw.Range, doc='Name, as a string, of the model which will be assessed')
@xw.arg('PF_Active_Energy_Net', np.array, doc='')
@xw.arg('PF_Reactive_Energy_Net', np.array, doc='')
@xw.ret(index=False, header=False, expand='down')
def PF(model_name, PF_Active_Energy_Net, PF_Reactive_Energy_Net):
    """Convenience function to assemble a dataframe for calculating performance ratio with unconstrained power"""

    if not isKoalaModelCached(model_name.name.name):
        generateModelGraph(model_name)

    inputs_for_PF = pd.DataFrame({'PF_Active_Energy_Net': PF_Active_Energy_Net, 'PF_Reactive_Energy_Net': PF_Reactive_Energy_Net})
    return EvaluateKoalaModel(model_name.name.name, inputs_for_PF)
