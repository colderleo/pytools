import pymysql, itertools, pyexcel_xls, warnings, re, datetime

from typing import List

def import_xls(xls_path, sheet_to_table, mysql_conf, db_columes_row=1, data_start_row=2):
    '''
        usage:
        target_xls = 'file_name.xlsx'
        sheet_to_table = [
            {"sheet_name": "sheel1", "db_table_name": "table1"},
            {"sheet_name": "sheel2", "db_table_name": "table2"},
        ]
        mysql_conf = {
            'NAME': 'test_db',
            'USER': 'root',
            'PASSWORD': '123456',
            'HOST': '127.0.0.1',
            'PORT': '3306', 
        }
        import_xls(target_xls, sheet_to_table, mysql_conf, db_columes_row=1, data_start_row=2)
    '''
    total_xls_data = pyexcel_xls.get_data(xls_path)
    DB = DBConn(mysql_conf, ignore_warning=False)
    
    for sheet_conf in sheet_to_table:
        sheet_name = sheet_conf["sheet_name"]
        import_data_to_db(total_xls_data[sheet_name], DB, sheet_conf['db_table_name'], sheet_name, \
            db_columes_row=db_columes_row, data_start_row=data_start_row)

    del DB


def import_data_to_db(sheet_data, DB, table_name, sheet_name, db_columes_row=1, data_start_row=2):
    db_names = sheet_data[db_columes_row]  # the row records db_column_name
    pure_data = sheet_data[data_start_row:]
    cursor = DB.dict_cursor

    for i, db_name in enumerate(db_names):
        if db_name.find(':') >= 0:
            col_setting = db_name.split(':')
            pure_db_name = col_setting[0]

            foreign_names = col_setting[1].split('.')
            foreign_table = foreign_names[0]
            foreign_context_col = foreign_names[1]
            
            foreign_key_col = col_setting[2]

            sql = 'select {},{} from {}'.format(foreign_context_col, foreign_key_col, foreign_table)
            cursor.execute(sql)
            records = cursor.fetchall()
            foreign_dict = {}
            for record in records:
                foreign_dict[record[foreign_context_col]] = record[foreign_key_col]

            db_names[i] = pure_db_name
            for row in pure_data:
                if row[i] in foreign_dict:
                    row[i] = foreign_dict[row[i]]
                elif row[i]:
                    print('Alert: {} foreign key not found'.format(row[i]))

    db_name_str = '({})'.format(','.join([val for val in db_names if val]))
    insert_data_strs = []
    for i, row in enumerate(pure_data):
        row_values = []
        for db_name, row_value in itertools.zip_longest(db_names, row):
            if db_name:
                # if type(row_values==str):
                if isinstance(row_value, int) or isinstance(row_value, float):
                    row_value = '{}'.format(row_value)
                elif row_value:
                    row_value = '"{}"'.format(row_value)
                else:
                    row_value = 'DEFAULT'
                row_values.append(row_value)
        insert_data_strs.append('({})'.format(','.join(row_values)))

    sql = 'insert ignore into {} {} VALUES {}'.format( \
        table_name, db_name_str, ','.join(insert_data_strs))
    res = cursor.execute(sql)
    DB.conn.commit()
    print(f'{sheet_name} import succeed {res}/{len(pure_data)}.')


class DBConn:
    def __init__(self, db_conf:dict, ignore_warning=False, create_db=False):
        '''
            db_conf is a dict like: {
                'NAME': 'test_db',
                'USER': 'root',
                'PASSWORD': '123456',
                'HOST': '127.0.0.1',
                'PORT': '3306',
            }
            db_conf['NAME'] can be a empty str.
            If create_db=True, will create db (if db not exists) with utf8.
        '''
        mysql_host = db_conf['HOST']
        mysql_user = db_conf['USER']
        mysql_passwd = db_conf['PASSWORD']
        mysql_port = int(db_conf['PORT'])
        mysql_dbname = db_conf['NAME']

        # connect db
        self.conn = pymysql.connect(host=mysql_host, user=mysql_user, passwd=mysql_passwd, port=mysql_port)
        self.dict_cursor = self.conn.cursor(cursor=pymysql.cursors.DictCursor)
        self.common_cursor = self.conn.cursor()

        if db_conf['NAME']:
            if create_db:
                self.common_cursor.execute(f"CREATE DATABASE IF NOT EXISTS {db_conf['NAME']} DEFAULT CHARACTER SET utf8 COLLATE utf8_general_ci")
            self.dict_cursor.execute("use " + mysql_dbname)  # use default database
            self.common_cursor.execute("use " + mysql_dbname)

        # remove 'ONLY_FULL_GROUP_BY' for sql_mode  
        sql = 'set @@global.sql_mode=`STRICT_TRANS_TABLES,NO_ZERO_IN_DATE,NO_ZERO_DATE,ERROR_FOR_DIVISION_BY_ZERO,NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION`'
        self.dict_cursor.execute(sql)
        self.common_cursor.execute(sql)

        if ignore_warning:
            warnings.filterwarnings("ignore",category=pymysql.Warning)


    def __del__(self):
        self.conn.close()
        self.dict_cursor.close()
        self.common_cursor.close()


def get_date_by_str(date_str:str):
    'convert a data-like string to date, such as: 2020-6-12, 2020.6.12, 2020.06.12, 20200612'
    spliter = re.sub(r'\d', '',date_str) # remove all number to get a spliter
    if not spliter:
        res_date = datetime.date(int(date_str[:4]), int(date_str[4:6]), int(date_str[6:]))
        return res_date
    spliter = spliter[0]
    
    ymd = date_str.split(spliter)[:3]
    res_date = datetime.date(int(ymd[0]), int(ymd[1]), int(ymd[2]))
    return res_date


def to_list(item):
    'if input item is a list, return it. if not, return [item]'
    if isinstance(item, list):
        return item
    return [item]


def date_ext(year:int, month:int, day:int):
    '''
        Create a datetime.date with extend (year,month,day), but the month and day can be any integer num:
            date_ext(2019,2,30) -> datetime.date(2019,3,2)
            date_ext(2019,2,0) -> datetime.date(2019,1,31)
            date_ext(2019,0,15) -> datetime.date(2018,12,15)
            date_ext(2019,-1,15) -> datetime.date(2018,11,15)

        some useful way to use:
        1. get each month's last day of year 2019:
            for i in range(2, 14):
                print(date_ext(2019, i, 0))
        2. get next month's 15 (no metter it will go next year or not):
            today = datetime.date.today()
            print(date_ext(today.year, today.month+1, 15))
    '''
    while month>12:
        year += 1
        month -= 12
    while month<1:
        month += 12
        year -= 1
    delta_days = datetime.timedelta(day-1)
    return datetime.date(year, month, 1)+delta_days


class TradingDayCalc:
    def __init__(self, holidays:List[datetime.date]):
        self.holidays = holidays

    def is_trading_day(self, the_date:datetime.date):
        'check if the_date is a trading day'
        if the_date in self.holidays:
            return False
        if the_date.weekday() in [5,6]:
            return False
        return True

    def holiday_go_before(self, the_date):
        'get the latest trading day on the_date or before the_date'
        while not self.is_trading_day(the_date):
            the_date += datetime.timedelta(-1)
        return the_date

    def holiday_go_after(self, the_date):
        'get the earlist trading day on the_date or after the_date'
        while not self.is_trading_day(the_date):
            the_date += datetime.timedelta(1)
        return the_date

