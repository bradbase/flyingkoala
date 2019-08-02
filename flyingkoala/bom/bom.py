from datetime import datetime

import xlwings as xw
import numpy as np
import pandas as pd

from flyingkoala import flyingkoala as fk

@xw.func
@xw.arg('file_name', 'raw', doc='Full path file name for the .axf file.')
@xw.ret(expand='table', ndim=2, index=None)
def load_observations_axf(file_name):
    """Function to extract the data table from a Bureau of Meteorology observations .axf"""

    with open(file_name, 'r') as file:
        data = False
        header = True
        df = None
        header_cells = None
        returnable = []

        for line in file:
            if line == '[$]\n':
                data = False

            if data:
                if header:
                    header_cells = line.split(',')
                    returnable.append(header_cells)
                    header = False
                else:
                    record_cells = line.split(',')
                    record_cells[2] = record_cells[2][1:-1]
                    record_cells[3] = record_cells[3][1:-1]
                    record_cells[5] = datetime.strptime(record_cells[5],'"%Y%m%d%H%M%S"')
                    record_cells[6] = datetime.strptime(record_cells[6],'"%Y%m%d%H%M%S"')
                    returnable.append(record_cells)

            if line == '[data]\n':
                data = True

    return returnable
