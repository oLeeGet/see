# Welcome to my libary
# 20220701
from pykeepass import PyKeePass
import datetime
import dateutil.relativedelta
import cx_Oracle
from cx_Oracle import Connection
import mysql.connector
import sqlite3
import pandas
import numpy
import traceback
import tempfile
import os

_KPBase="D:/Google/0_Documents/1.5_Key/db.kdbx"
_KPKey="D:/Google/0_Documents/1.5_Key/db.keyx"

#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
# DataBase
#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
#Get DataBase Connection
def get_con(_Type:str="Oracle",_title:str="oLeeGetOracleDB") -> Connection:
    """
    Get DataBase Connection

    Option(1): Oracle (Default)
    Option(2): Mysql
    Option(3): Sqlite
    """
    if _Type == "Oracle":
        _TnsName="h3g_online"
        kp = PyKeePass(_KPBase, keyfile=_KPKey)
        entry = kp.find_entries(title=_title, first=True)
        connection = cx_Oracle.connect(
            entry.username, 
            entry.password, _TnsName
            )
    elif _Type == "Mysql":
        connection = mysql.connector.connect(
            database="DW3G",
            host="localhost",
            user="fp_admin",
            password="qwerty"
            )
    elif _Type == "Sqlite":
        connection = sqlite3.connect(
            database='c:/sqlite/db2.db'
            )
    return connection

#Sql to DataFrame
def run_sql(conn:Connection, sql:str) -> pandas.DataFrame:
    return pandas.read_sql_query(sql, con=conn)

#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
# DataTime
#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

def get_dt():
    return datetime.datetime.today().strftime('%Y%m%d_%H%M%S')

def get_last_day_of_month(_yyyymm:str) -> str:
    """
    (Sample)
        get_last_day_of_month('202202'), will retun: 0228
    """
    
    date_object = datetime.datetime.strptime(_yyyymm, '%Y%m').date()

    # this will never fail
    # get close to the end of the month for any day, and add 4 days 'over'
    next_month = date_object.replace(day=28) + datetime.timedelta(days=4)
    
    # subtract the number of remaining 'overage' days to get last day of current month, or said programattically said, the previous day of the first of next month
    eom = (next_month - datetime.timedelta(days=next_month.day)).strftime('%m%d')

    return eom

#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
# Quick Work
#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
# Open file in Excel
def open_file(file:str):
    os.system("start EXCEL.EXE " + file)

# Gen file path
class save_as:
    def __init__(self, Name:str="tmp", To:str="tmp"):
        """
        Name:The File Name Prefix
            "Sales Report" = "Sales Report_{yyyymmdd}.xlsx"
        To:Where you want save to 
            "tmp" = system defaut temp path
            "my" = my D:/temp
        """
        self.fprefix = Name
        if To == "tmp":
            self.tmp_dir = tempfile.gettempdir()
        elif To == "my":
            self.tmp_dir = "D:/temp"
    
    def type(self, file_type:str="xlsx") -> str:
        """
        Save to Excel? CSV? TXT? etc...
        """
        ftype_dist = {
            'excel': 'xlsx',
            'xlsx': 'xlsx',
            'csv': 'csv',
            'txt': 'txt',
            'json': 'json',
            'html': 'html',
            'web': 'html',
            'www': 'html'
            }
        ftype = ftype_dist.get(file_type,"xlsx")
        return self.tmp_dir + "/" + self.fprefix + "_" + get_dt() + "." + ftype

# Gen file pw
def gen_file_pw(_pw='rt') -> str:
    """
    Default password prefix is 'rt'
    (Sample)
        Today: 1 Oct 2022
        Return: rt2022q4
    """
    x = (datetime.date.today() + dateutil.relativedelta.relativedelta(days=-1)).strftime('%m')
    _quarter = {
        '01': '1',
        '02': '1',
        '03': '1',
        '04': '2',
        '05': '2',
        '06': '2',
        '07': '3',
        '08': '3',
        '09': '3',
        '10': '4',
        '11': '4',
        '12': '4'
    }
    _qr = _quarter.get(x,'9') # 9 will be returned default if _mth is not found
    _file_pw = _pw + (datetime.date.today() + dateutil.relativedelta.relativedelta(days=-1)).strftime('%Y') + 'q' + _qr

    return _file_pw

# Check order details into Excel file
def check_order(order_id:str) -> str:
    file = gen_file_path().type()

    try:
        writer = pandas.ExcelWriter(file) #, engine="xlsxwriter")

        dic1 = {
            "Name": ["Alan","April","John","Sarah"],
            "Age": [30,24,28,26],
            "Sex": ["M","F","M","F"]
        }
        df1 = pandas.DataFrame(dic1)

        dic2 = {
            "AAA": numpy.arange(1,21),
            "BBB": numpy.arange(21,41),
            "CCC": numpy.arange(41,61),
            "DDD": numpy.arange(61,81),
            "EEE": numpy.arange(81,101)
        }
        df2 = pandas.DataFrame(dic2)

        df1.to_excel(writer, sheet_name="RAW1")
        df2.to_excel(writer, sheet_name="RAW2")

        print("The result saved in " + file)

        return file
    except:
        print("error")
        traceback.print_exc()
    finally:
        writer.save()


# Testing
class Check_Order:

    @classmethod
    def to_string(str, format):
        order, account, mobile = format.split(",")
        return str(order, account, mobile)

    def __init__(self, order_id, account_no, mobile_no):
        self.order = order_id
        self.account = account_no
        self.mobile = mobile_no

    #show account details
    def show_ac_mob(self):
        self.account = 12345
        return f"Account: {self.account}, Mobile: {self.mobile}"

    #show like what
    def show_like(self, something):
        return f"he like {something} ..."

class Dog:
    def __init__(self,name):
        self.__name = name
        self.__tricks = []
    
    def add_trick(self, trick):
        self.__tricks.append(trick)
    def print_trick(self):
        print("My name: " + self.__name + str(self.__tricks))