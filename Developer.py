import math
import pyodbc
import requests
import pandas as pd

from bs4 import BeautifulSoup
from datetime import datetime

from sqlalchemy import exc
from sqlalchemy.sql import text
from sqlalchemy import create_engine

## 取得系統現在時間
def get_datetime():
    time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    return time

def connect_sql(sql_db=None):  
    # if sql_db is None:
    #     sql_db = r'Origin'
    server = r""
    database = r""
    username = r""
    password = r""
    engine = create_engine('mssql+pyodbc://'+username+':'+password+'@'+server+':1433/'+database+'?driver=SQL+Server+Native+Client+11.0')
    return engine

def query2sql(sql):
    engine = connect_sql()
    engine.execute(text(sql).execution_options(autocommit=True))
    # engine.execute(sql)
    engine.dispose()
def insertMost(year, host, unit, name, report, Time, Total, ck, ek, ca, ea, page):
    name = str.replace(name ,'\'' ,'"')
    ck = str.replace(ck ,'\'' ,'"')
    ek = str.replace(ek ,'\'' ,'"')
    ca = str.replace(ca ,'\'' ,'"')
    ea = str.replace(ea ,'\'' ,'"')
    
    query = """
        declare @Year nvarchar(10) = N'{0}'
        declare @Host nvarchar(100) = N'{1}'
        declare @Unit nvarchar(100) = N'{2}'
        declare @Project nvarchar(max) = N'{3}'
        declare @Report nvarchar(100) = N'{4}'
        declare @Duration nvarchar(100) = N'{5}'
        declare @Total nvarchar(100) = N'{6}'
        declare @Ck nvarchar(max) = N'{7}'
        declare @Ek nvarchar(max) = N'{8}'
        declare @Ca nvarchar(max) = N'{9}'
        declare @Ea nvarchar(max) = N'{10}'
        declare @Page nvarchar(10) = N'{11}'
        exec xp_insertMOST @Year, @Host, @Unit, @Project, @Report, @Duration, @Total, @Ck, @Ek, @Ca, @Ea, @Page
         
    """.format(year, host, unit, name, report, Time, Total, ck, ek, ca, ea, page)
    try:
        query2sql(query)
        print(name)
    except:
        insertexception(query)

def insertLink(link, Year, Host, Pjt):
    Pjt = str.replace(Pjt, '\'', '"')
    query = """
        declare @Link nvarchar(1000) = N'{0}'
        declare @Year nvarchar(10) = N'{1}'
        declare @Host nvarchar(1000) = N'{2}'
        declare @Name nvarchar(max) = N'{3}'

        exec xp_insertMOSTLink @Year, @Host, @Name, @Link
    """.format(link, Year, Host, Pjt)
    try:
        query2sql(query)
    except:
        insertexception(query)

def insertexception(query):
    f = open('log.txt','a+',encoding = 'UTF-8')
    f.write(str(query)+'\n')
    f.close

def xpath2text(driver, xpath):
    text = driver.find_element_by_xpath(xpath).text
    return text

def xpath2onclick(driver, xpath):
    text = driver.find_element_by_xpath(xpath).get_attribute('onclick')
    return text

def get_Report(driver, i):
    Report_str = '//*[@id="wUctlAwardQueryPage_grdResult_ctl%02d_lblAWARD_REPORT_STATUSc"]' % i
    try :
        text = xpath2text(driver, Report_str)
    except:
        text = '已上傳'
    return text

def get_nYear(driver, Year, Host, Pjt, i):
    nYear_str = '//*[@id="wUctlAwardQueryPage_grdResult_ctl%02d_lnkAWARD_TOT_AUD_AMT"]' % i
    try:
        text = xpath2onclick(driver, nYear_str).split('\'')[1]
        insertLink(text, Year, Host, Pjt)
    except:
        text = 'NO nYear.'

def get_CkLink(driver, Year, Host, Pjt, i):
    CkLink_str = '//*[@id="wUctlAwardQueryPage_grdResult_ctl%02d_lnkWattBRIEF_CHI_S"]' % i
    try:
        text = xpath2onclick(driver, CkLink_str).split('\'')[1]
        insertLink(text, Year, Host, Pjt)
    except:
        text = 'NO CkLink.'

def get_EkLink(driver, Year, Host, Pjt, i):
    EkLink_str = '//*[@id="wUctlAwardQueryPage_grdResult_ctl%02d_lnkWattBRIEF_ENG_S"]' % i
    try:
        text = xpath2onclick(driver, EkLink_str).split('\'')[1]
        insertLink(text, Year, Host, Pjt)
    except:
        text = 'NO EkLink.'