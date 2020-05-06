
import unittest
import logging

import xlwings as xw
from flyingkoala import FlyingKoala
from pandas import DataFrame
from numpy import array
from numpy.testing import assert_array_equal

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

        self.workbook_name = r'growing_degrees_day.xlsm'

        if len(xw.apps) == 0:
            raise "We need an Excel workbook open for this unit test."

        self.my_fk = FlyingKoala(self.workbook_name, load_koala=True)
        self.my_fk.reload_koala('')
        self.equation_name = xw.Range('Equation_1')

        if self.equation_name not in self.my_fk.koala_models.keys():
            model = None
            wb = xw.books[self.workbook_name]
            wb.activate()
            for name in wb.names:
                self.my_fk.load_model(self.equation_name)
                if self.equation_name == name.name:
                    model = xw.Range(self.equation_name)
                    self.my_fk.generate_model_graph(model)

            if model is None:
                return 'Model "%s" has not been loaded into cache, if named range exists check spelling.' % self.equation_name


    def test_Equation_1(self):
        """First type of test for Equation_1"""

        xw.books[self.workbook_name].sheets['Growing Degree Day'].activate()

        goal = xw.books[self.workbook_name].sheets['Growing Degree Day'].range(xw.Range('D2'), xw.Range('D6')).options(array).value

        tmin = xw.books[self.workbook_name].sheets['Growing Degree Day'].range(xw.Range('B2'), xw.Range('B6')).options(array).value
        tmax = xw.books[self.workbook_name].sheets['Growing Degree Day'].range(xw.Range('C2'), xw.Range('C6')).options(array).value
        inputs_for_DegreeDay = DataFrame({'T_min': tmin, 'T_max': tmax})
        result = self.my_fk.evaluate_koala_model('Equation_1', inputs_for_DegreeDay).to_numpy()

        assert_array_equal(goal, result)


    def test_VBA_Equation_1(self):
        """
        The function definition being called;

        Function VBA_Equation_1(T_min As Double, T_max As Double) As Double
            VBA_Equation_1 = Application.WorksheetFunction.Max(((T_max + T_min) / 2) - 10, 0)
        End Function
        """

        goal = 20

        vba_equation_1 = xw.books[self.workbook_name].macro('VBA_Equation_1')
        result = vba_equation_1(20.0, 40.0)

        self.assertEqual(goal, result)
