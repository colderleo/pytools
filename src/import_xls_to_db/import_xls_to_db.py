import pymysql, itertools, pyexcel_xls, warnings


CONFIG = {
    'target_xls' : 'import_xls_to_db_test.xlsx',
    'sheet_to_table' : [
        {"sheet_name": "学生", "db_table_name": "student"},
        {"sheet_name": "分数", "db_table_name": "score"},
    ],
    'mysql' : {
        'NAME': 'test_db',
        'USER': 'root',
        'PASSWORD': '123456',
        'HOST': '127.0.0.1',
        'PORT': '3306', 
    },
    'extra' : {
        'xls_format' : {
            "db_name_row": 1,
            "start_row": 2,
        }
    }
}



def main():
    import_xls(CONFIG)


def import_xls(config):
    DB = DBConn(config['mysql'], ignore_warning=False)
    sheet_to_table_conf = config['sheet_to_table']
    total_xls_data = pyexcel_xls.get_data(config['target_xls'])
    
    for sheet_conf in sheet_to_table_conf:
        sheet_name = sheet_conf["sheet_name"]
        import_data_into_db(total_xls_data[sheet_name], config['extra']['xls_format'], DB, sheet_conf['db_table_name'], sheet_name)

    del DB


def import_data_into_db(data, xls_format, DB, tar_table, info_symbol):
    cursor = DB.dict_cursor
    start_row = xls_format['start_row']
    pure_data = data[start_row:]
    db_names = data[xls_format['db_name_row']]

    # process data with foreign key
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
                    print('Alert: {} 外键未找到'.format(row[i]))

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
        tar_table, db_name_str, ','.join(insert_data_strs))
    res = cursor.execute(sql)
    DB.conn.commit()
    print('{} 导入完成，成功导入{}/{}条'.format(info_symbol, res, len(pure_data)))


class DBConn:
    def __init__(self, db_conf:dict, ignore_warning=False):

        mysql_host = db_conf['HOST']
        mysql_user = db_conf['USER']
        mysql_passwd = db_conf['PASSWORD']
        mysql_port = int(db_conf['PORT'])
        mysql_dbname = db_conf['NAME']

        # connect db
        self.conn = pymysql.connect(host=mysql_host, user=mysql_user, passwd=mysql_passwd, port=mysql_port)
        self.dict_cursor = self.conn.cursor(cursor=pymysql.cursors.DictCursor)
        self.common_cursor = self.conn.cursor()

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




if __name__ == "__main__":
    main()
