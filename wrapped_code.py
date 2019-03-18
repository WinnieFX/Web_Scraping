from selenium import webdriver
import time
import os
import pandas as pd
import glob 
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import xml.etree.ElementTree as ET
import xml
import pyodbc
import logging

import win32com.client as win32
#-- Create log file
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

formatter = logging.Formatter('%(asctime)s:%(name)s:%(message)s')

file_handler = logging.FileHandler('cenace_gob.log')
file_handler.setFormatter(formatter)

if (logger.hasHandlers()):
    logger.handlers.clear()

logger.addHandler(file_handler)

#-- get the current directory and set it as the Working directory

cwd = os.getcwd()
os.chdir(cwd)

def num_to_ith(num):
    """1 becomes 1st, 2 becomes 2nd, etc."""
    value = str(num)    
    last_digit = value[-1]
    if len(value) > 1 and value[-2] == '1': return value +'th'
    if last_digit == '1': return value + 'st'
    if last_digit == '2': return value + 'nd'
    if last_digit == '3': return value + 'rd'
    return value + 'th'

def cenacedmd(Nyear,Nday):
    """
    Nyear: how many years, e.g., 1 for the latest 1 year, 2 for the latest 2 years, etc.
    Nday: how many days, e.g., 1 for today, 2 for the latest 2 days, etc.

    """
    #-- Firefox configuration
    download_dir = cwd
    fp = webdriver.FirefoxProfile()
    fp.set_preference("browser.download.folderList",2)
    fp.set_preference("browser.download.dir", download_dir)
    fp.set_preference("browser.download.manager.showWhenStarting", False)
    fp.set_preference("browser.helperApps.neverAsk.openFile", "application/octet-stream")
    fp.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/octet-stream")

    driver = webdriver.Firefox(firefox_profile=fp)
    
    demand_url = f'http://www.cenace.gob.mx/SIM/VISTA/REPORTES/H_RepCantAsignadas.aspx?N=59&opc=divCssCantAsig&site=Cantidades%20asignadas/MDA/De%20Energ%C3%ADa%20El%C3%A9ctrica%20por%20Zona%20de%20Carga&tipoArch=C&tipoUni=SIN&tipo=De%20Energ%C3%ADa%20El%C3%A9ctrica%20por%20Zona%20de%20Carga&nombrenodop=MDA' 
    driver.get(demand_url)
    time.sleep(3)
    driver.refresh()
    time.sleep(2)

    #-- After opening the website, use xpath to click the latest N years 
   
    error1 = {}
    for i in range(1,Nyear+1) : 
        Yorder = num_to_ith(i)
        yrs_ago = datetime.now() - relativedelta(years=(i-1))
        yrs = str(yrs_ago.year)
        try:
            html = driver.find_element_by_xpath("""/html/body/form/div[4]/div[1]/div/div[3]/div[3]/div/table/tbody/tr/td[1]/div/ul/li[%d]/div/span[3]""" %i)
            html.click()
            logger.info('On the website' + ': ' +'The '+ yrs + ' year is selected')
        except Exception as e:
            error1[i] = 'On the website' + ': ' + 'The ' + yrs +' year does not exist' 
            logger.info('On the website' + ': ' + 'The ' + yrs +' year does not exist')
    time.sleep(3)
    #-- Define the Latest N days
    now = datetime.today()
    #print(str(now)[:10])    
   
    alldates = []
    for j in range(0,Nday):
        date_ago = now - timedelta(days=j)
        date_ago_new = str(date_ago)[:10]
        #print(date_ago_new)
        alldates.append(date_ago_new)

    #-- On the website, after hiting the year bottoms, select the files for wanted dates to click---
    error2 = {}
    for i in range(0,len(alldates)):
        strdate = str(alldates[i])
        try:
            elems = driver.find_elements_by_xpath("""//a[contains(@id,'CSV') and contains(@href, "%s") ]""" % strdate)
            for elem in elems:
                #print(elem)
                elem.click()
                time.sleep(2)
                #elemhref = elem.get_attribute("href")
                #print(elemhref)
            logger.info('The files for ' + strdate + ' are downloaded')
        except:
            error2[i] = 'There is no file for ' + strdate + ', on the website'
            logger.info('There is no file for ' + strdate + ', on the website')


    #-- close driver
    time.sleep(3)       
    driver.quit()
    #path = r"path = r'D:\Users\fanxin\Downloads\web_scraping"

    #-- read saved csv files
    dic_result = {}
    ls_date = {}
    error3 = {}
    info = {}

    for i in ['BCS','BCA','SIN']:
        dic_result[i] = pd.DataFrame()
        ls_date[i] = []
        try:
            for name in glob.glob( 'Can'+'*' + i + '*.csv'):
                df = pd.read_csv(name, skiprows = 7,index_col=None, header=0)
                df = df.iloc[:,[0,1,4]]
                df.columns = ['Demand_Point','Hour','Value']
                date = name.split(' ')[-6]
                ls_date[i].append(date)
                sp_date = date.split('-')
                year = sp_date[0]
                month = sp_date[1]
                day = sp_date[2]
                df['Year'] = year
                df['Month'] = month
                df['Day'] = day
                df['Source'] = 'CENACE'
                df['Date_of_Entry'] = str(datetime.now())
                df['UserName'] = 'NULL'
                df['Batch'] = 'NULL'
                df['Type'] = 'NULL'
                dic_result[i] = dic_result[i].append(df)
                os.remove(name) 
                conn = pyodbc.connect(driver='{SQL Server Native Client 10.0}',server='houcardo2',database='Cardo',Trusted_Connection='Yes')
                cursor = conn.cursor()
                cursor.execute("DELETE FROM Demand_Hourly_wf_tmp WHERE Source = 'CENACE' and Year = ? and Month = ? and Day = ? ",(str(year),str(month),str(day)))
                cursor.commit()
            
            dic_result[i].to_csv(i+'_'+'test.csv')
            
            start_date = min(ls_date[i])
            end_date = max(ls_date[i])
            logger.info('Data for ' + i + ' from '+ start_date +' to ' + end_date +' are saved into the local folder')
            info[i] = 'Data for ' + i + ' from '+ start_date +' to ' + end_date +' are saved into the local folder' 
        except:
            error3[i] = 'No Data is available for ' + i
            logger.info('No Data is available for ' + i + ' ')
    
    for i in ['BCS','BCA','SIN']:
        dic_result[i] = dic_result[i][['Demand_Point','Year','Month','Day','Hour','Value','Source','Date_of_Entry','UserName','Batch','Type']]
        
        #---
        demand_matrix = dic_result[i].as_matrix()
        root = ET.Element("root")
        for j in range(demand_matrix.shape[0]):
            demand_records = ET.SubElement(root, "demand",
                                          dict(
                                                 Demand_Point = str(demand_matrix[j,0]),
                                                 Year = str(demand_matrix[j,1]),
                                                 Month = str(demand_matrix[j,2]),
                                                 Day = str(demand_matrix[j,3]),
                                                 Hour = str(demand_matrix[j,4]),
                                                 Value = str(demand_matrix[j,5]),
                                                 Source = str(demand_matrix[j,6]),
                                                 Date_of_Entry = str(demand_matrix[j,7]),
                                                 UserName = str(demand_matrix[j,8]),
                                                 Batch = str(demand_matrix[j,9]),
                                                 Type = str(demand_matrix[j,10])
                        
                                                  ))
            
        xmlstr_demand = ""
        xmlstr_demand = ET.tostring(root, encoding='us-ascii', method='xml').decode('utf-8')
        sql_demand_hourly = """
        DECLARE @idoc INT, @doc VARCHAR(MAX);
        SET @doc = '""" + xmlstr_demand + """'
        EXEC sp_xml_preparedocument @idoc OUTPUT, @doc;
        INSERT INTO Demand_Hourly_wf_tmp ( Demand_Point,Year,Month,Day,Hour,Value,Source,Date_of_Entry,UserName,Batch,Type )
        SELECT Demand_Point,Year,Month,Day,Hour,Value,Source,Date_of_Entry,UserName,Batch,Type
        FROM OPENXML (@idoc, '/root/demand')
        WITH (
               Demand_Point varchar(100),
               Year float,
               Month float,
               Day float,
               Hour float,
               Value float,
               Source varchar(50),
               Date_of_Entry datetime,
               UserName varchar(50),
               Batch varchar(50),
               Type varchar(12)
               );
        """
        with pyodbc.connect(driver='{SQL Server Native Client 10.0}',server='houcardo2',database='Cardo',Trusted_Connection='Yes') as conn:
            cursor = conn.cursor()
            cursor.execute(sql_demand_hourly)
            conn.commit()

    #-- Combine error messages
    error_list = []
    error_list =list(error1.values())+list(error2.values())+list(error3.values())
    if len(error_list)>0:
        try:
            for i in range(len(error_list)):
                error_msg = "\n".join(str(e) for e in error_list)
        except:
            print('No Error')
            
    
    #-- information messages
    information=[]
    information = list(info.values())
    if len(information)>0:
        try:
            for i in range(len(information)):
                info_msg = "\n".join(str(e) for e in information)
        except:
            print('No Data')

    
    #-- Open outlook and sen
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    #-- get email adresses the messages will be sent to
    
    #--- open a text file written with email addresses 
    mailto = open('Send_Email_to_(Email_Address_List).txt','r+')
    address = mailto.read()
    address = address.replace('\n',';')
    mailto.close()
    #--- assign these addresses to parameter 'mail.To'
    mail.To = address
    
    mail.Subject = 'Updates of Scriping Data from CENACE Demand'
    """
    Set the format of mail
    1 - Plain Text
    2 - HTML
    3 - Rich Text
    """
    if len(error_list)>0 and len(information)>0:
 
        mail.Body = 'Good Morning!'+ '\n'+ '\n'+ 'Error Message: ' + '\n'+ error_msg + '\n' + 'Information for Data Downloaded: '+ '\n'+ info_msg+ '\n'+ '\n'+ 'Regards,' + '\n' + 'Winnie'
        
    elif len(error_list)==0 and len(information)>0:
        mail.Body = 'Good Morning!'+ '\n'+ '\n'+ 'Information for Data Downloaded: '+ '\n'+ info_msg+ '\n'+ '\n'+ 'Regards,' + '\n' + 'Winnie'
        
    elif len(error_list)>0 and len(information)==0:
        mail.Body = 'Good Morning!'+ '\n'+ '\n'+ 'Error Message: ' + '\n'+ error_msg + '\n'+ '\n'+ 'Regards,' + '\n' + 'Winnie'
    else:
        mail.Body = 'Good Morning!'+ '\n'+ '\n'+ 'Warning: ' + '\n' + 'Nothing from ' + demand_url +'\n'+ '\n'+ 'Regards,' + '\n' + 'Winnie'
    mail.BodyFormat = 1
    mail.Send()
    
    cursor.close()

    return error1,error2,error3,info


if __name__ == '__main__':
    
    results = cenacedmd(1,1)
    
print('All jobs done')

