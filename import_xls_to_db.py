
from tools import import_xls

def main():
    target_xls = 'import_xls_to_db_test.xlsx'
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
    import_xls(target_xls, sheet_to_table, mysql_conf)


if __name__ == "__main__":
    main()
