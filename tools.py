# pylint: disable=import-error

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
    import pyexcel_xls, itertools
    total_xls_data = pyexcel_xls.get_data(xls_path)
    DB = DBConn(mysql_conf, ignore_warning=False)
    
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


    for sheet_conf in sheet_to_table:
        sheet_name = sheet_conf["sheet_name"]
        import_data_to_db(total_xls_data[sheet_name], DB, sheet_conf['db_table_name'], sheet_name, \
            db_columes_row=db_columes_row, data_start_row=data_start_row)

    del DB



def convert_xls_col_chars_to_indexx(chars):
    'convert A->1, Z->26, AA->27, AB->28, AZZ->1378. indexx starts from 1'
    ret = 0
    for char in chars:
        ret *= 26
        value = ord(char) - ord('A') + 1
        ret += value
    return ret

def convert_xls_col_indexx_to_chars(indexx:int):
    'convert 1->A, 26->Z, 17->AA, 28->AB, 1378->AZZ. indexx starts from 1'
    ret = ''
    ord_zero = ord('A') - 1
    radix = ord('Z') - ord_zero
    while indexx > radix:
        rest = indexx % radix
        if (rest==0):
            ret += 'Z'
            indexx = indexx // radix - 1
        else:
            ret += chr(indexx % radix + ord_zero)
            indexx = indexx // radix
    if indexx > 0:
        ret += chr(indexx + ord_zero)
    return ret[::-1]


import pymysql, warnings
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

        warnings.filterwarnings("ignore",category=pymysql.Warning) #ignore warnings for already exists table or db, or doesn't exist 
        if db_conf['NAME']:
            if create_db:
                self.common_cursor.execute(f"CREATE DATABASE IF NOT EXISTS {db_conf['NAME']} DEFAULT CHARACTER SET utf8 COLLATE utf8_general_ci")
            self.dict_cursor.execute("use " + mysql_dbname)  # use default database
            self.common_cursor.execute("use " + mysql_dbname)

        # remove 'ONLY_FULL_GROUP_BY' for sql_mode  
        sql = 'set @@global.sql_mode=`STRICT_TRANS_TABLES,NO_ZERO_IN_DATE,NO_ZERO_DATE,ERROR_FOR_DIVISION_BY_ZERO,NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION`'
        self.dict_cursor.execute(sql)
        self.common_cursor.execute(sql)

        if not ignore_warning:
            warnings.resetwarnings()

    def __del__(self):
        self.conn.close()
        self.dict_cursor.close()
        self.common_cursor.close()


def get_date_by_str(date_str:str):
    import re, datetime
    'convert a string to date. support many kinds of format like: 2020-6-12, 2020.6.12, 2020.06.12, 20200612'
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
            date_ext(2019,2,30)  -> datetime.date(2019,3,2)
            date_ext(2019,2,0)   -> datetime.date(2019,1,31)
            date_ext(2019,0,15)  -> datetime.date(2018,12,15)
            date_ext(2019,-1,15) -> datetime.date(2018,11,15) 

        some useful way to use:
        1. get each month's last day of year 2019:
            for i in range(2, 14):
                print(date_ext(2019, i, 0))
        2. get next month's 15 (don't worry it will go next year):
            today = datetime.date.today()
            print(date_ext(today.year, today.month+1, 15))
    '''
    import datetime
    while month>12:
        year += 1
        month -= 12
    while month<1:
        month += 12
        year -= 1
    delta_days = datetime.timedelta(day-1)
    return datetime.date(year, month, 1)+delta_days


import datetime
class TradingDayCalc:
    from typing import List
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


import warnings, pymysql
class FilterWarning(object):
    '''
        with FilterWarning():
            DB.connect()
    '''
    def __init__(self):
        pass

    def __enter__(self, category=pymysql.Warning):
        # print('FilterWarning: filter pymysql warnings')
        warnings.filterwarnings("ignore",category=category)
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        # print('FilterWarning: reste warnings')
        warnings.resetwarnings()
        return True


class PrintException(object):
    '''
        with PrintException():
            a = 1/0
    '''
    def __init__(self, print_trace=True):
        self.print_trace = print_trace
        pass

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, exc_traceback_obj):
        print(f'PrintException: {exc_value}')
        if self.print_trace:
            import traceback
            traceback.print_exception(exc_type, exc_value, exc_traceback_obj)
        return True


def wrapperTpl(func):
    from functools import wraps
    @wraps(func)
    def inner(*args, **kwargs):
        '''在执行目标函数之前要执行的内容'''
        ret = func(*args, **kwargs)
        '''在执行目标函数之后要执行的内容'''
        return ret
    return inner



def dict_recursive_update(default, custom):
    '''Copy from github: https://github.com/Maples7/dict-recursive-update
    Return a dict merged from default and custom
    >>> recursive_update('a', 'b')
        Traceback (most recent call last):
            ...
        TypeError: Params of recursive_update should be dicts
    '''
    if not isinstance(default, dict) or not isinstance(custom, dict):
        raise TypeError('Params of recursive_update should be dicts')

    for key in custom:
        if isinstance(custom[key], dict) and isinstance(
                default.get(key), dict):
            default[key] = dict_recursive_update(default[key], custom[key])
        else:
            default[key] = custom[key]

    return default


def load_json(file_path:str):
    import json
    with open(file_path,'r',encoding='utf-8') as f:
        json_var = json.load(f)
        return json_var

def save_json(json_var, file_path:str):
    import json
    '''attention: will rewrite the file of file_path'''
    with open(file_path, 'w',encoding='utf-8') as f:
        json.dump(json_var, f, ensure_ascii=False, indent=4)

def random_in(i:int):
    import random
    return random.randint(1,i)==1


def get_xls_cell_date(cell):
    import xlrd
    if cell.ctype == 3: 
        ms_date_number = cell.value
        year, month, day, _hour, _minute, _second = xlrd.xldate_as_tuple(ms_date_number, xlrd.Book.datemode)
        res_date = datetime.date(year, month, day) 
        return res_date
    else:
        date_str = str(cell.value)
        return get_date_by_str(date_str)


def print_obj(*args):
    for obj in args: 
        if type(obj) == str:
            print(obj)
        else:
            print('\n'.join(['%s:%s' % item for item in obj.__dict__.items()]) )


def code_response(code=1, msg='', data:dict={}):
    from django.http.response import JsonResponse
    ret = {
        'rescode': code,
        'resmsg': msg,
    }
    ret.update(data)
    return JsonResponse(ret)


def get_timestamp(t:datetime.datetime=None):
    import datetime, time
    if not t:
        t = datetime.datetime.now()
    return int(time.mktime(t.timetuple()))


def load_json_var(json_str:str, key=None):
    import json
    with PrintException:
        ret = json.loads(json_str)
        if ret and key:
            return ret.get(key)
        return ret
    return None


def generate_jwt_token(openid:str='undefined_wx_openid', encode_to_str=True):
    import jwt # pip install pyjwt
    JWT_SECRET = 'dkdll893hj938h42h829h'
    EXPIRE_SECONDS = '7000'
    payload = {
        'uid': 'uid_abc',
        'expire_time': get_timestamp() + EXPIRE_SECONDS,
    }
    token = jwt.encode(payload, JWT_SECRET, algorithm='HS256')  # decoded = jwt.decode(token, secret, algorithms='HS256')
    if encode_to_str:
        token = str(token, encoding='utf-8')
    return token


def parse_jwt_token(token, key=None):
    '''
        if verify passed, return payload
        if key, return payload[key]
        else, return None
    '''
    import jwt # pip install pyjwt
    try:
        JWT_SECRET = 'dkdll893hj938h42h829h'
        payload = jwt.decode(token, JWT_SECRET, algorithms='HS256')
        cur_timestamp = get_timestamp()
        if cur_timestamp > payload['expire_time']:
            raise Exception('login expired')
        if key:
            return payload.get(key)
        else:
            return payload
    except Exception as e:
        print(f'verify token failed: {e}')
    return None

