"""
FlyingKoala provides the ability to dynamically define Python calculations from Excel formulas enabling you to replace the Excel calculation engine with Python with a genuine minimum of Python code.

The main benefit of replacing existing formulas with FyingKoala is that you get faster calculation, less need to get your hands dirty and can unit test your formulas. It speeds up scenario analysis and model development

Two users;
Those who want to easily add python functions to excel
Those who want to speed up their existing crazy big models

"""

import setuptools

with open("README.md", "r") as fh:
    long_description = fh.read()

setuptools.setup(
    name="flyingkoala",
    version="0.0.10b2",
    author="Bradley van Ree",
    author_email="flyingkoala@bradbase.net",
    description="Integration of xlwings and xlcalculator with MS Excel plug-in",
    long_description=long_description,
    long_description_content_type="text/markdown",
    keywords=['xls',
        'excel',
        'spreadsheet',
        'workbook',
        'vba',
        'macro',
        'data analysis',
        'analysis'
        'reading excel',
        'excel formula',
        'excel formulas',
        'excel equations',
        'excel equation',
        'formula',
        'formulas',
        'equation',
        'equations',
        'pandas',
        'harvest',
        'timeseries',
        'time series',
        'energy',
        'accounting',
        'horticulture',
        'research',
        'visualization',
        'scenario analysis',
        'modelling',
        'model',
        'unit testing',
        'testing',
        'audit'],
    url="https://github.com/bradbase/flyingkoala",
    packages=setuptools.find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: GNU General Public License v3 (GPLv3)",
        'Operating System :: Microsoft :: Windows',
        'Operating System :: MacOS :: MacOS X',
    ],
    install_requires=[
            'xlwings >= 0.15.8',
            'numpy >= 1.15.0',
            'pandas >= 0.25.0',
            'xlfunctions <= 0.0.3b0',
            'xlcalculator <= 0.0.5b0',
        ]
)
