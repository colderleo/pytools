# pyexcel-xls很好用， xls_data = pyexcel_xls.get_data('test.xlsx')

def find_row_col(data_2d, target, startswith=False, index_0_based=True, col_chars_format=False):
    '''
        index_0_based=True: index starts from 0.
        find target's row and col in sheet. 
        data_2d is a 2d array, or iterable by rows, cols.
        startswith: find cell which is str and startswith(target)
        col_chars_format: return xls col name lick A,B...Z,AA,AB
        if not found, return None, None
    '''
    def find():
        for i, row in enumerate(data_2d):
            for j, value in enumerate(row):
                if value==target:
                    return (i, j)
                if startswith and isinstance(value, str) and value.startswith(target):
                    return (i, j)
        return None, None

    row, col = find()
    if row is not None: # found
        if not index_0_based:
            row += 1
            col += 1
        if col_chars_format:
            col_chars = xls_column_num_to_letter(col+1)
            col = col_chars
    return row, col

        



def get_xlrd_cell_date(cell):
    import xlrd, datetime
    from tools_common import get_date_by_str
    if cell.ctype == 3: 
        ms_date_number = cell.value
        year, month, day, _hour, _minute, _second = xlrd.xldate_as_tuple(ms_date_number, xlrd.Book.datemode)
        res_date = datetime.date(year, month, day) 
        return res_date
    else:
        date_str = str(cell.value)
        return get_date_by_str(date_str)



def xls_column_letter_to_num(chars):
    '''
        convert A->1, Z->26, AA->27, AB->28, AZZ->1378. 
        the same as: openpyxl.utils.column_index_from_string
        num starts from 1.
    '''
    ret = 0
    for char in chars:
        ret *= 26
        value = ord(char) - ord('A') + 1
        ret += value
    return ret


def xls_column_num_to_letter(num:int):
    '''
        convert 1->A, 26->Z, 17->AA, 28->AB, 1378->AZZ. num starts from 1
        the same as: from openpyxl.utils.get_column_letter
    '''
    ret = ''
    ord_zero = ord('A') - 1
    radix = ord('Z') - ord_zero
    while num > radix:
        rest = num % radix
        if (rest==0):
            ret += 'Z'
            num = num // radix - 1
        else:
            ret += chr(num % radix + ord_zero)
            num = num // radix
    if num > 0:
        ret += chr(num + ord_zero)
    return ret[::-1]





