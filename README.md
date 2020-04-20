# leo-pytools
some useful tools and cases outside standard python library.

## import xls to database
import xls data to mysql database, with supporting for foeign key.

### test
#### create database
Connect to mysql and execute the following sql to create a test database.
```sql
CREATE DATABASE IF NOT EXISTS `test_db`;

CREATE TABLE `test_db`.`student`(
	`id` int PRIMARY KEY,
	`name` varchar(20) NOT NULL
);

CREATE TABLE `test_db`.`score`(
    `id` int PRIMARY KEY AUTO_INCREMENT,
	`student_id` int,
	`subject` varchar(20) NOT NULL,
    `score` float NOT NULL
);
```

#### install python packages
install the following packages in your python environment: pymysql, pyexcel_xls

#### run
run `python import_xls_to_db.py`,and the data in xls will be imported to mysql.


#### set a foreign key
`student_id:student.name:id` means find from foreign table.colume 'student.name' and get it's id, then put the id in this table's 'student_id' colume.

