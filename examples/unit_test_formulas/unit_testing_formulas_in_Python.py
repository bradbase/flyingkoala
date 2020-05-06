
import unittest
import logging

import xlwings as xw
from flyingkoala import FlyingKoala
from pandas import DataFrame
from pandas import Series
from numpy import array
from pandas.util.testing import assert_series_equal

logging.basicConfig(level=logging.ERROR)

class Test_equation_1(unittest.TestCase):
    """Unit testing Equation_1 in Excel

    Using;
    - named ranges to discover the address of the formula
    - Excel as the calculation engine

    This approach requires;
    - Workbook to be tested needs to be open. In this case it's example.xlsm from the example directory in FlyingKoala


    python -m unittest discover -p "unit_test*.py"
    """


    def setUp(self):

        self.workbook_name = r'./examples/unit_test_formulas/growing_degrees_day.xlsm'

        if len(xw.apps) != 0:
            raise "We want all Excel workbooks closed for this unit test."

        self.my_fk = FlyingKoala(self.workbook_name, load_koala=True)
        self.my_fk.reload_koala('')
        self.equation_name = xw.Range('Equation_1')

        if self.equation_name not in self.my_fk.koala_models.keys():
            self.my_fk.load_model(self.equation_name)
            if self.equation_name == name.name:
                model = xw.Range(self.equation_name)
                self.my_fk.generate_model_graph(model)


    def test_Equation_1(self):
        """First type of test for Equation_1"""

        goal = Series([0.0, 0.0, 0.0, 0.0, 0.0, 5.0, 10.0, 15.0, 20.0])

        tmin = array([-20, -15, -10, -5, 0, 5, 10, 15, 20])
        tmax = array([0, 5, 10, 15, 20, 25, 30, 35, 40])
        inputs_for_DegreeDay = DataFrame({'T_min': tmin, 'T_max': tmax})
        result = self.my_fk.evaluate_koala_model('Equation_1', inputs_for_DegreeDay)

        assert_series_equal(goal, result)
