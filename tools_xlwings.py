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


def insert_value_by_mark(xls_filename, row_mark, col_mark, value, sheet_name=None):
    'inster value buy finding row_mark and col_mark.'
    import pyexcel_xls
    from tools_xls import find_row_col

    xls_total_data = pyexcel_xls.get_data(xls_filename)
    if sheet_name is None:
        xls_data = xls_total_data[next(iter(xls_total_data))]
    else:
        xls_data = xls_total_data[sheet_name]

    tar_rowi, _col = find_row_col(xls_data, row_mark)
    if tar_rowi is None:
        raise Exception(f"'{xls_filename}'内找不到{row_mark}")
    _row, tar_coli = find_row_col(xls_data, col_mark)
    if tar_coli is None:
        raise Exception(f"'{xls_filename}'内找不到{col_mark}")
        
    # workbook = xlwings.Book(xls_filename)
    workbook = app.books.open(xls_filename, update_links=True)
    if sheet_name is None:
        sheet:xlwings.Sheet = workbook.sheets[0]
    else:
        sheet:xlwings.Sheet = workbook.sheets[sheet_name]
    tar_cell = sheet[tar_rowi, tar_coli]
    tar_cell.value = value
    workbook.save()
    print(f'    {xls_filename} - {row_mark} - {col_mark} - {value} 写入成功')


def get_or_create_xlwings_app():
    app = xlwings.apps.active #优先使用现有的excel app实例，避免创建新的app导致冲突
    if not app:
        app = xlwings.App(visible=True, add_book=False)
        print('create new excel app')
    app.visible = True
    app.display_alerts = True #如果设置为False 文件修改后关闭会不提示保存
    return app


def xlwings_demo():
    app = get_or_create_xlwings_app()

    wb = app.books.open('xxx.xlsx') #open an existing xls. if it's already opened, will connect, not reopen.
    wb = xlwings.Book()  # this will create a new workbook

    sheet:xlwings.Sheet = wb.sheets[0]  # sheet = wb.sheets['Sheet1']
    max_row = sheet.used_range.last_cell.row
    max_col = sheet.used_range.last_cell.column

    # 下拉自动填充
    fill_ref_range = sheet[0, 3:max_col] #填充参考区域
    fill_full_range = sheet[0:max_row, 0:max_col] #要填充的所有区域
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


