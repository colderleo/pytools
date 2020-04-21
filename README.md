# leo-pytools
some useful tools and cases outside standard python library.

## import xls to database
import xls data to mysql database, with supporting for foeign key. Test it in the following way.

#### create database
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

#### install python packages
install the following packages in your python environment: pymysql, pyexcel_xls

#### run
run `python import_xls_to_db.py`,and the data in xls will be imported to mysql.


#### xls data format
the second line is the colume name in database.  if the colume's content is a foreign key, you can set it as following:

`student_id:student.name:id` means find from foreign table.colume 'student.name' and get it's id, then put the id in this table's 'student_id' colume.

1. student sheet:

| 学号 | 姓名 | 备注  |
| ---- | ---- | ----- |
| id   | name |       |
| 1001 | 张三 | 备注1 |
| 1002 | 李四 | 备注2 |

2. score sheet:

| 学生姓名                   | 学科    | 分数  |
| -------------------------- | ------- | ----- |
| student_id:student.name:id | subject | score |
| 张三                       | 语文    | 98    |
| 李四                       | 数学    | 86    |





