import io
import sys
import unittest

import xlwings as xw
import flyingkoala
from flyingkoala import *

sys.setrecursionlimit(3000)

workbook_name = r'example.xlsm'

class Test_equation_1(unittest.TestCase):
    """Unit testing Equation_1 in Excel

    Using;
    - named ranges to discover the address of the formula
    - Excel as the calculation engine

    This approach requires;
    - Workbook to be tested needs to be open. In this case it's example.xlsm from the example directory in FlyingKoala
    """

    equation_name = 'Equation_1'

    books = xw.books

    equation_1 = None
    selected_book = None
    input_addresses = {}

    # find the equation, and its address
    for book in books:
        if book.name == workbook_name:
            selected_book = book

            for named_range in book.names:
                if named_range.name == equation_name:
                    equation_1 = named_range

            # discover the input terms for the equation
            inputs = parse_model(equation_1.refers_to_range)

            # Now find the input term addresses.
            for term in inputs:
                for named_range in book.names:
                    if term == named_range.name:
                        input_addresses[term] = named_range.refers_to_range

    def test_1(self):
        """First type of test for Equation_1"""

        # define test case inputs
        case_00 = {'T_base': 10, 'T_min': 10, 'T_max': 20, 'result': 5.0}

        # set the input values with the test case
        for term in self.input_addresses:
            self.input_addresses[term].value = case_00[term]

        # Do a re-calc and stay focussed!
        xw.App(add_book=False, visible=False).calculate()
        self.selected_book.activate()

        # test the result of the calculation
        self.assertEqual(case_00['result'], self.equation_1.refers_to_range.value)
