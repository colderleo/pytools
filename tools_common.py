# py--lint: disable=import-error



def send_email(sender:str, receivers:list, msg_title, msg_body, smtp_server:str, password:str, 
            cc_emails:list=None, sender_name:str='',
            attachment_filepath=None, attachment_name=None, 
            receivers_can_see_eachother=False, print_ret=False) -> bool:
        '''
            receivers_can_see_eachother: default False, each receiver can only see himself. 
            cc_emails: cc_emails can always be seen by all receivers.
            attachment_name: default set to the initial filename
            sender_name: the name shown in email beside sender email address, like: Mike<123456@qq.com>

            useage:
            tools_common.send_email(sender='test@qq.com', receivers=['receiver1@163.com'], \
                msg_title='标题A', msg_body='正文B', smtp_server='smtp.qq.com', password='123456',
                cc_emails=['cc1@qq.com'], attachment_filepath='test.txt')
        '''


        import smtplib, os
        from email.mime.text import MIMEText
        from email.utils import formataddr
        from email.mime.multipart import MIMEMultipart
        from email.mime.application import MIMEApplication

        receivers = to_list(receivers)

        msg = MIMEMultipart()
        msg['From'] = formataddr([sender_name, sender])
        msg['Subject'] = msg_title
        if cc_emails:
            msg['Cc'] = ','.join(to_list(cc_emails))
        if receivers_can_see_eachother or (len(receivers)==1):
            msg["To"] = ','.join(receivers)
            
        msg.attach(MIMEText(msg_body, 'plain', 'utf-8'))

        if attachment_filepath:
            with open(attachment_filepath, 'rb') as f:
                if attachment_name is None:
                    _folder_path, attachment_name = os.path.split(attachment_filepath);
                part = MIMEApplication(f.read())
                part.add_header('Content-Disposition', 'attachment', filename=attachment_name)
                msg.attach(part)

        try:
            smtp = smtplib.SMTP()
            smtp.connect(smtp_server)
            smtp.login(sender, password)
            if cc_emails:
                for cc in cc_emails:
                    if cc not in receivers:
                        receivers.append(cc)
            vaild_receivers = []
            for r in receivers:
                if r:
                    vaild_receivers.append(r)
            smtp.sendmail(sender, vaild_receivers, msg.as_string())
            if(print_ret):
                print('邮件发送成功')
            return True
        except smtplib.SMTPException as e:
            if(print_ret):
                print('邮件发送失败: ' + str(e))
            return False
        smtp.quit()


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
        if the_date in self.holidays:
            return False
        if the_date.weekday() in [5,6]:
            return False
        return True

    def holiday_go_before(self, the_date):
        ret = the_date + datetime.timedelta(0)   # copy input
        while not self.is_trading_day(ret):
            ret += datetime.timedelta(-1)
        return ret

    def holiday_go_after(self, the_date):
        ret = the_date + datetime.timedelta(0)  # copy input
        while not self.is_trading_day(ret):
            ret += datetime.timedelta(1)
        return ret

    def n_trading_days_later(self, start_day:datetime.date, n:int):
        '''
            usage: 1 trading day later(T1): n_trading_day_later(today, 1)
            no metter if start_day is trading day.
        '''
        ret = start_day + datetime.timedelta(0)  # copy input
        for _i in range(n):
            ret += datetime.timedelta(1)
            ret = self.holiday_go_after(ret)
        return ret



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



def print_obj(*args):
    for obj in args: 
        if type(obj) == str:
            print(obj)
        else:
            print('\n'.join(['%s:%s' % item for item in obj.__dict__.items()]) )


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

