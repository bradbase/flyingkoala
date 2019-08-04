.. _worked_example_timeseries:

Worked Example - Time Series
============================

The worked example for the timeseries module demonstrates the use of FlyingKoala User Defined Functions for time series transformation.

1. The FlyingKoala Add-In
-------------------------

The FlyingKoala Add-In assists users to manage the cache which holds "models" (/equation systems). This is what it looks like.

.. image:: /images/addin.png
  :alt: Add-in menu

Excel file name
^^^^^^^^^^^^^^^

This is the name of the workbook file. This setting is expected to auto-fill.

In the future we expect to be able to load models from other Excel files which is why this configuration is here.

.. image:: /images/addin_excel_file_name.png
  :width: 400
  :alt: Add-in menu excel file name



  For doing some time series transformation you would have a module that would look like this:

  .. code-block:: Python

    import xlwings as xw
    from flyingkoala import flyingkoala
    from flyingkoala.timeseries import *
