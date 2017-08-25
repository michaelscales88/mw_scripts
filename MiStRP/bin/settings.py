import os
import pyexcel as pe

from os import path


def valid_input(column_position, row, ws, input_type):
    # Utility function for validating cells from each ws
    if input_type == 'N':
        try:
            return_value = int(ws['%s%d' % (column_position, row)].value)
        except TypeError:
            return_value = 0
        except ValueError:
            try:
                return_value = ws['%s%d' % (column_position, row)].value
                h, m, s = [int(float(i)) for i in return_value.split(':')]
            except TypeError:
                return_value = 0
            except ValueError:
                return_value = 0
            else:
                return_value = (3600 * int(h)) + (60 * int(m)) + int(s)
    elif input_type == 'B':
        return_value = ws['%s%d' % (column_position, row)].value
    elif input_type == 'S':
        return_value = str(ws['%s%d' % (column_position, row)].value)
    else:
        raise ValueError("Invalid input type in valid_input")
    return return_value


SELF_PATH = os.path.dirname(path.dirname(path.abspath(__file__)))
constants = pe.get_dict(file_name='%s/bin/config.xlsx' % SELF_PATH, name_columns_by_row=0)
index_dict = {}
for index, item in enumerate(constants['Constant']):
    index_dict[item] = index
GO_PIC = constants['Argument'][index_dict['GO_PIC']]
ACCUM_PIC = constants['Argument'][index_dict['ACCUM_PIC']]
QUIT_PIC = constants['Argument'][index_dict['QUIT_PIC']]
SEARCH_PIC = constants['Argument'][index_dict['SEARCH_PIC']]
SETTINGS_PIC = constants['Argument'][index_dict['SETTINGS_PIC']]
CALL_SLA_ARG = constants['Argument'][index_dict['CALL_SLA_ARG']]
ACCUM_ARG = constants['Argument'][index_dict['ACCUM_ARG']]
SPREADSHEET_VIEWER_FILE_TEMPLATE = constants['Argument'][index_dict['SPREADSHEET_VIEWER_FILE_TEMPLATE']]
ACCUM_VIEWER_FILE = constants['Argument'][index_dict['ACCUM_VIEWER_FILE']]
TEST_SPREADSHEET = constants['Argument'][index_dict['TEST_SPREADSHEET']]

