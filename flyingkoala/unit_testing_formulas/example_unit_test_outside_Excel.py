import io
import sys
import unittest

import xlwings as xw
import flyingkoala
from flyingkoala import *

sys.setrecursionlimit(3000)

workbook_path = 'C:\\Users\\yourself\\Documents\\Python\\example\\'
workbook_name = r'example.xlsm'

class Test_equation_1(unittest.TestCase):
    """Unit testing Equation_1 in Python

    Using;
    - named ranges to discover the address of the formula
    - Python as the calculation engine

    This approach requires;
    - Workbook to be tested needs to be open. In this case it's example.xlsm from the example directory in FlyingKoala++

    ++This needs to be changed so that FlyingKoala can use Koala to figure out the model for itself
    """

    equation_name = 'Equation_1'

    books = xw.books
    workbook = reload_koala('%s%s' % (workbook_path, workbook_name), ignore_sheets=['Raw Data'])

    equation_1 = None
    selected_book = None

    # find the equation, and its address
    for book in books:
        if book.name == workbook_name:
            selected_book = book

            for named_range in book.names:
                if named_range.name == equation_name:
                    equation_1 = named_range.refers_to_range
                    # parse the equation into Python
                    generate_model_graph(equation_1)

    def test_1(self):
        """First type of test for Equation_1"""

        # define test case inputs
        case_00 = {'terms' : {'T_base': 10, 'T_min': 10, 'T_max': 20}, 'result' : 5.0}

        # Do a calc
        result = flyingkoala.evaluate_koala_model_row(self.equation_name, case_00['terms'], flyingkoala.koala_models[self.equation_name])

        # test the result of the calculation
        self.assertEqual(case_00['result'], result)
