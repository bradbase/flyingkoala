.. _installation:

Install the library
===================

The easiest way to install FlyingKoala is via pip::

    pip install FlyingKoala

Dependencies
------------

* ``xlwings``, ``xlcalculator``, ``pandas``, ``numpy``

Optional Dependencies
---------------------

The FlyingKoala supplied modules which have no extra dependencies.

* Horticulture
* Time Series

These packages are not required but highly recommended as they play very nicely with xlwings.

* Matplotlib
* Pillow/PIL

Install the add-in
==================

The trouble with installing this one is the element of "it depends". The punchline is the add-in needs to be placed where your add-ins go.

Until we can arrange a script that figures it out for us we will need to do it by hand.

Copy the addin\\flyingkoala.xlam to;

Here:
C:\\Users\\username\\AppData\\Roaming\\Microsoft\\AddIns

Or here:
C:\\Users\\username\\AppData\\Roaming\\Microsoft\\Excel\\XLSTART

Sometimes there's an XLSTART in your home directory.

It could well be somewhere else... (especially if you're on a Mac)

If you got through the add-in install and are following the example bouncing ball, next up is :ref:`use`.
