from selenium import webdriver
from selenium.webdriver.support.ui import Select
from Developer import *
import time

driver = webdriver.Chrome()
driver.get("https://was.most.gov.tw/WAS2/Award/AsAwardMultiQuery.aspx")
driver.find_element_by_xpath('//*[@id="dtlItem_ctl00_btnItem"]').click()
Select(driver.find_element_by_name('wUctlAwardQueryPage$repQuery$ctl01$ddlYRst')).select_by_value("78")
Select(driver.find_element_by_name('wUctlAwardQueryPage$repQuery$ctl01$ddlYRend')).select_by_value("107")
Select(driver.find_element_by_name('wUctlAwardQueryPage$ddlPageSize')).select_by_value("200")
driver.find_element_by_xpath('//*[@id="wUctlAwardQueryPage_btnQuery"]').click()

engine = connect_sql()
sql = """ select Page from RawDB..CheckCrawl where Status = 0 order by Page desc """
df = pd.read_sql(sql, engine)
for i in range(0, len(df)):

    page = df.iloc[i]['Page']
    Select(driver.find_element_by_name('wUctlAwardQueryPage$grdResult$ctl203$ddlPage')).select_by_value(str(page))
    time.sleep(10)

    for i in range(2, 201+1, 1):
        Year_str = '//*[@id="wUctlAwardQueryPage_grdResult"]/tbody/tr[{0}]/td[1]'.format(str(i))
        Host_str = '//*[@id="wUctlAwardQueryPage_grdResult"]/tbody/tr[{0}]/td[2]'.format(str(i))
        Unit_str = '//*[@id="wUctlAwardQueryPage_grdResult"]/tbody/tr[{0}]/td[3]'.format(str(i))
        Pjt_str = '//*[@id="wUctlAwardQueryPage_grdResult_ctl%02d_lblAWARD_PLAN_CHI_DESCc"]' % i
        Time_str = '//*[@id="wUctlAwardQueryPage_grdResult_ctl%02d_lblAWARD_ST_ENDc"]' % i
        Total_str = '//*[@id="wUctlAwardQueryPage_grdResult_ctl%02d_lblAWARD_TOT_AUD_AMTc"]' % i
        Ck_str = '//*[@id="wUctlAwardQueryPage_grdResult_ctl%02d_lblKEYS_CHIc"]' % i
        Ek_str = '//*[@id="wUctlAwardQueryPage_grdResult_ctl%02d_lblKEYS_ENGc"]' % i
        Ca_str = '//*[@id="wUctlAwardQueryPage_grdResult_ctl%02d_lblBRIEF_CHI_Sc"]' % i
        Ea_str = '//*[@id="wUctlAwardQueryPage_grdResult_ctl%02d_lblBRIEF_ENG_Sc"]' % i    

        Year = xpath2text(driver, Year_str)
        Host = xpath2text(driver, Host_str)
        Unit = xpath2text(driver, Unit_str)
        Pjt = xpath2text(driver, Pjt_str)
        Report = get_Report(driver, i)
        Time = xpath2text(driver, Time_str)
        Total = xpath2text(driver, Total_str)
        Ck = xpath2text(driver, Ck_str)
        Ek = xpath2text(driver, Ek_str)
        Ca = xpath2text(driver, Ca_str)
        Ea = xpath2text(driver, Ea_str)
        insertMost(Year, Host, Unit, Pjt, Report, Time, Total, Ck, Ek, Ca, Ea, page)
        get_nYear(driver, Year, Host, Pjt, i)
        get_CkLink(driver, Year, Host, Pjt, i)
        get_EkLink(driver, Year, Host, Pjt, i)
    # driver.close()


## 判斷有無繳交報告 done!
## 換頁
## 寫入資料庫 done!
## 時間分 ST ET
## 紀錄多年期 done!
## 紀錄"詳" done!
