import xlwings
def find_xlwings_cell(rng:xlwings.Range, target, startswith=False):
    'cell is a 1x1 Range. 非常慢'
    for row in rng.rows:
        for cell in row:
            if (cell.value == target) or (startswith and cell.value.startswith(target)):
                return cell


def get_xlwings_col(cell:xlwings.Range):
    'cell is a 1x1 Range'
    sheet = cell.sheet
    for col in sheet.used_range.columns:
        if col.column == cell.column:
            return col


def get_xlwings_total_range(sheet:xlwings.Sheet) -> xlwings.Range:
    'return range from A1 to all used range'
    return sheet[:sheet.used_range.last_cell.row, :sheet.used_range.last_cell.column]


def xlwings_demo():
    import xlwings

    app = xlwings.apps.active
    if not app:
        app = xlwings.App(visible=True, add_book=False)
        print('create new excel app')
    app.visible = True
    app.display_alerts = False
    app.books.open('xxx.xlsx') #这样会使用现有的excel app实例，不会创建新的导致冲突
    
    wb = xlwings.Book()  # this will create a new workbook
    # wb = xlwings.Book('FileName.xlsx')  # connect to an existing file in the current working directory
    sheet:xlwings.Sheet = wb.sheets[0]  # sheet = wb.sheets['Sheet1']
    max_row = sheet.used_range.last_cell.row
    max_col = sheet.used_range.last_cell.column

    fill_ref_range = sheet[0, 3:max_col]
    fill_full_range = sheet[0:max_row, 0:max_col]
    # sheet.range('B103:AQ103').api.AutoFill(sheet.range('B103:AQ104').api,0)
    fill_ref_range.api.AutoFill(fill_full_range.api, 0)

    # an useful way to deal data:
    sheet = wb.sheets[0]
    product_name_col_index = find_xlwings_cell(sheet.used_range, 'product_name').colunm - 1
    equity_col_index = find_xlwings_cell(sheet.used_range, 'equity').colunm - 1
    equity_dict = {}
    for row in sheet.used_range.rows:
        product_name = sheet[row.row-1, product_name_col_index].value
        equity = sheet[row.row-1, equity_col_index].value
        if isinstance(product_name, str) and (type(equity) in [float, int]):
            equity_dict[product_name] = equity
    print(equity_dict)


