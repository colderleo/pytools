# pytools
some useful python tools.

## import xls to database

#### useage

```python
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
import_xls(target_xls, sheet_to_table, mysql_conf, db_columes_row=1, data_start_row=2)
```

#### xls data format - support foreign key

the **db_columes_row** of xls sheet indicate the colume name of database table.  if the colume's content is a foreign key, you can set it as following:

`student_id:student.name:id`, which means find the record from foreign table 'student.name' and get it's id, then put the id in this table's 'student_id' colume.

1. student sheet:

   

| 学号 | 姓名 | 备注    |
| ---- | ---- | ------- |
| id   | name |         |
| 1001 | stu1 | remark1 |
| 1002 | stu2 | remark2 |

2. score sheet:

   first line is ignored when import, and you can write the header whatever you want.

| Student Name               | Important Subject | Score |
| -------------------------- | ----------------- | ----- |
| student_id:student.name:id | subject           | score |
| Stu1                       | subject1          | 98    |
| Stu2                       | subject2          | 86    |

#### test

import xls data to mysql database, with supporting for foeign key. Test it in the following way.

##### create database
Connect to mysql and execute the following sql to create a test database.
```sql
CREATE DATABASE IF NOT EXISTS `test_db`;

CREATE TABLE `test_db`.`student`(
	`id` int PRIMARY KEY,
	`name` varchar(20) NOT NULL
);

CREATE TABLE `test_db`.`score`(
	`student_id` int,
	`subject` varchar(20) NOT NULL,
    `score` float NOT NULL,
	primary key (`student_id`, `subject`)
);
```

##### install depending python packages
install the following packages in your python environment: pymysql, pyexcel_xls

##### run
run `python import_xls_to_db.py`,and the data in xls will be imported to mysql.





