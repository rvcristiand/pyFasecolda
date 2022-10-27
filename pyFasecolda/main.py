import os

from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

import xlwings as xw

import time
import pandas as pd
import numpy as np

# dictionaries
actividades_economicas_dict = {
    'TODAS': 1,
    '1000010': 2,
    '1701001': 3,
    '1702001': 4,
    '1713001': 5,
    '1721001': 6,
    '1722001': 7,
    '1723001': 8,
    '1724001': 9,
    '1729001': 10,
    '1731001': 11,
    '1741101': 12,
    '1741201': 13,
    '1741301': 14,
    '1741401': 15,
    '1742101': 16,
    '1743001': 17,
    '1749101': 18,
    '1749901': 19,
    '2000010': 20,
    '2711101': 21,
    '2711201': 22,
    '2711301': 23,
    '2712101': 24,
    '2712201': 25,
    '2712301': 26,
    '2725001': 27,
    '2731002': 28,
    '2742102': 29,
    '2749201': 30,
    '2749301': 31,
    '2749401': 32,
    '2749501': 33,
    '2749901': 34,
    '3000010': 35,
    '3731002': 36,
    '3732001': 37,
    '3742102': 38,
    '3749101': 39,
    '3749501': 40,
    '4000010': 41,
    '4711102': 42,
    '4711202': 43,
    '4711302': 44,
    '4712102': 45,
    '4749202': 46,
    '4749302': 47,
    '4749402': 48,
    '4749502': 49,
    '5000010': 50,
    '5701001': 51,
    '5712201': 52,
    '5742101': 53,
    '5742201': 54,
    '5749203': 55,
    '5749303': 56
}

year_dict = {
    '2009': 1,
    '2010': 2,
    '2011': 3,
    '2012': 4,
    '2013': 5,
    '2014': 6,
    '2015': 7,
    '2016': 8,
    '2017': 9,
    '2018': 10,
    '2019': 11,
    '2020': 12,
    '2021': 13,
    '2022': 14
}


def join_xls_files():
    """Join xls files.

    Join xls files stored in '.\Clean_dataset' in a csv file
    """
    # files xls' folder path
    list_files=os.listdir('.\Clean_dataset')

    # list xls files only
    list_excel=[l for l in list_files if l.endswith('.xls')]

    # fill the dataframe
    df_all=pd.DataFrame()
    for f in list_excel:
        path=os.path.join('.\Clean_dataset', f)
        print(path)
        info=pd.read_excel(path,  skiprows=[0,1], nrows=8, usecols=[2,6],index_col=0,names=['index','value'])
        info2=pd.read_excel(path,  skiprows=[0,1], nrows=8, usecols=[12,16],index_col=0,names=['index','value'])
        info=info.append(info2)
        info=info.dropna().T
        df=pd.read_excel(path,  skiprows=np.arange(0,14), usecols=[1,3,8,9,17,18,19,21,23,24,26,27],index_col=0)
        df.loc[:,'año']=info['Año'][0]
        df.loc[:,'mes']=info['Mes'][0]
        df.loc[:,'departamento']=info['Departamento'][0]
        df_all=df_all.append(df.loc['TOTAL',:])

    df_all = df_all.replace(np.nan, 0).reset_index(drop=True)
    df_all.to_csv(os.path.join(os.getcwd(), 'Processed_dataset', 'fasecolda_dataset.csv'), index=False)


def fix_files():
    """Fix xls damaged files.

    Some downloaded reports are damaged and pandas can't process them. Excel
    can fix them, changed the damaged cells value to zero.

    The raw xls files are stored in './Raw_dataset' folder and the fixed
    files are stored at
    """
    # open Excel
    app = xw.App(visible=False)

    # do for xls files inside Raw_dataset folder
    for root, dirs, files in os.walk('.\Raw_dataset'):
        for xls_filename in files:
            if xls_filename.endswith('.xls'):
                # xls filepath
                filepath = os.path.join(root, xls_filename)
                print(f'file: {xls_filename}', end=' ')

                # open workbook
                wb = xw.Book(filepath)
                sh = wb.sheets[0]

                # used range
                used_range = sh.used_range
                no_rows = used_range.rows.count
                no_cols = used_range.columns.count

                cell = used_range[0, 0]

                # search for damage cells
                for i in range(no_rows):
                    for j in range(no_cols):
                        if cell.offset(i, j).value in [
                            -2146826281,
                            -2146826246,
                            -2146826259,
                            -2146826288,
                            -2146826252,
                            -2146826265,
                            -2146826273
                        ]:
                            # fix damage cell
                            cell.offset(i, j).value = 0

                # save fixed xls file
                wb.save(os.path.join(os.getcwd(), 'Clean_dataset', xls_filename))
                wb.close()

                print('changed')

    app.kill()


def download_reports(actividad_economica, year):
    """Download reports from Sistema General de Riesgos Laborales FASECOLDA.

    Download reports for all months and all departamentos from FASECOLDA's page
    web Sistema General de Riesgos Laborales, for Sector economico Inmobiliario.

    The downloaded files are stored at './Raw_dataset' folder.

    Args:
        actividad_economica: actividad economica
        year: year
    """
    # web explorer options
    firefoxOptions = Options()
    firefoxOptions.add_argument('--start-maximized')
    firefoxOptions.add_argument('--disable-extensions')

    # change download directory
    firefoxOptions.set_preference('browser.download.folderList', 2)
    firefoxOptions.set_preference("browser.download.dir", os.path.join(os.getcwd(), "Raw_dataset"))

    # create a firefox instance
    driver = webdriver.Firefox(executable_path="./geckodriver", options=firefoxOptions)

    # select the URL
    driver.get('https://sistemas.fasecolda.com/rldatos/Reportes/xCompania.aspx')

    # select year
    Icon_year = driver.find_element(By.ID, 'ctl00_ContentPlaceHolder1_rcbYears_Input').click()
    time.sleep(1)
    Year=driver.find_element(By.XPATH, f'//*[@id="ctl00_ContentPlaceHolder1_rcbYears_DropDown"]/div/ul/li[{year_dict[year]}]').click()
    time.sleep(1)
    
    # sector economico
    Icon_Sec=driver.find_element(By.ID, 'ctl00_ContentPlaceHolder1_rcbActEconomica_Input').click()
    time.sleep(0.5)
    # select inmobiliario
    Sec=driver.find_element(By.XPATH, '//*//*[@id="ctl00_ContentPlaceHolder1_rcbActEconomica_DropDown"]/div/ul/li[11]').click()
    time.sleep(0.5)

    # select actividad economica
    Icon_Act=driver.find_element(By.ID, 'ctl00_ContentPlaceHolder1_rcbSectEconomico_Input').click()
    time.sleep(0.5)
    Act=driver.find_element(By.XPATH, f'//*[@id="ctl00_ContentPlaceHolder1_rcbSectEconomico_DropDown"]/div/ul/li[{actividades_economicas_dict[actividad_economica]}]').click()
    time.sleep(0.5)

    for m in range(2,14):
        # select month
        Icon_Month=driver.find_element(By.ID, 'ctl00_ContentPlaceHolder1_rcbMeses_Input').click()
        time.sleep(1)
        Month=driver.find_element(By.XPATH, f'/html/body/form/div[1]/div/div/ul/li[{m}]').click()
        time.sleep(1)

        for d in range (2,35):
            # departamento
            Icon_Dpto=driver.find_element(By.ID, 'ctl00_ContentPlaceHolder1_rcbDeptos_Input').click()
            time.sleep(0.5)

            # municipios
            Dpto=driver.find_element(By.XPATH, f'//*[@id="ctl00_ContentPlaceHolder1_rcbDeptos_DropDown"]/div/ul/li[{d}]').click()
            time.sleep(0.5)

            # consultar
            Icon_submit=driver.find_element(By.ID, 'ctl00_ContentPlaceHolder1_Button1').click()
            time.sleep(0.5)

            # exportar como xls
            path_xls=driver.find_element(By.ID, 'ctl00_ContentPlaceHolder1_repVieCompania_ReportToolbar_ExportGr_FormatList_DropDownList')
            drop_xls =Select(path_xls)
            drop_xls.select_by_value('XLS')
            time.sleep(0.5)

            # descargar
            Icon_submit=driver.find_element(By.ID, 'ctl00_ContentPlaceHolder1_repVieCompania_ReportToolbar_ExportGr_Export').click()
            time.sleep(0.5)

        print('Done m:', m)

    driver.quit()


if __name__ == '__main__':
    join_xls_files()
    # fix_files()
    # download_2022()
