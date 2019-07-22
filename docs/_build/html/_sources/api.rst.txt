.. _api:

API
===

User Defined Functions
----------------------
These are the User Defined Functions UDFs that come with FlyingKoala.

energy
^^^^^^

A module with energy related functions.

energy
######

* SOLARNOON(latitude, longitude, times)

Uses PVLib to determine solar noon for a given latitude and longitude. Returns a dynamic array.

horticulture
^^^^^^^^^^^^

A module with horticulture related functions.

horticulture
############

* DegreesDay(model, T_min, T_max)

Uses Koala2 to reference an Excel formula definition for a degrees day calculation. Returns a numeric value.

* DegreesDayDynamicArray(model, T_min, T_max)

Uses Koala2 to reference an Excel formula definition for a degrees day calculation. Returns a dynamic array.

timeseries
^^^^^^^^^^

A module with time series related functions.

timeseries
############

* KEEPRECORDS(times, inputs, window=5)

Keeps records at an offset determined by window. Returns a Dynamic Array.

* RESAMPLEMINS(times, inputs, window=5, operation='mean')

Performs a look-ahead re-sample of size window with stated operation on a time series for values in inputs and does not return the time index. Returns a Dynamic Array.

* RESAMPLEMINSWITHINDEX(times, inputs, window=5, operation='mean')

Performs a look-ahead re-sample average of size window on a time series for values in inputs and returns with the time index. Returns a Dynamic Array.

* TIMEISBETWEEN(keys, below, above, periods)

Decides if a time is between certain range of a given time. Returns boolean.

* TIMESERIESWINDOWAVERAGE(times, inputs, window=5)

Performs a look-ahead moving average of size window on a time series for values in inputs. Returns a Dynamic Array.

VBA Macros
----------
These are the RunPython macros that come with FlyingKoala.

They don't work "out of the box" yet -- I've not written the plug-in but would work if you wrote your own RunPython calls.


Accounting
^^^^^^^^^^

A module with accounting related functions.

Harvest
#######

* get_harvest_invoices()

Queries Harvest for all invoices. Creates a worksheet called "Harvest Invoices" for them.

* get_harvest_time_entries()

Queries Harvest for all time entries. created a worksheet called "Harvest Time Entries" for them.

Utils
#####

* delete_columns(sheet_to_delete_from, columns_to_keep, sheet_rename, delete=True)

Deletes in place all columns from a given worksheet keeping only those named by column letter in columns_to_keep. Delete flag inverts the selection.

* keep_columns(sheet_to_delete_from, columns_to_delete, sheet_rename)

Convenience function. Same as delete_columns except defaults delete to False.

* make_employee_pivot_table()

Takes a table called "Employee Hours" and creates a pivot table.
