import time
import os
import glob
import pandas as pd 
import xlwings as xw

from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


def sfDownload(export_url, export_name):
    
    download_dir = f"{os.getcwd()}"
    export_name = f"{export_name}.csv"
    edge_options = Options()
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "profile.default_content_setting_values.automatic_downloads": 1
    }
    edge_options.add_experimental_option("prefs", prefs)
    driver = webdriver.Edge(options=edge_options)

    driver.get(export_url)
    WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Next']"))).click()

    time.sleep(10)
    downloaded = max(glob.glob(os.path.join(download_dir, "*")), key=os.path.getctime)

    if downloaded:
        new_path = os.path.join(download_dir, export_name)
        if os.path.exists(new_path): os.remove(new_path)
        os.rename(downloaded, new_path)
        print(f"Renamed to {new_path}")
    else:
        print("Download failed or timed out")
    
    driver.quit()

def delCw(sheet):
    last_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
    start_delete = last_row - 4
    sheet.api.Rows(f"{start_delete}:{last_row}").Delete()

    
#sfDownload(metroE, "MetroE")
metroE = pd.read_csv('MetroE.csv',encoding='windows-1252')


#mainwb = app.books.open('Open Project IRR Tool 1.1.xlsm')
# Tk().withdraw()
# file_path = askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xlsm;*.xlsb")])
# time.sleep(2)

mainwb = xw.apps.active.books.active
time.sleep(5)

print("MetroE")
mainwb.sheets["MetroE_Data"].clear()
mainwb.sheets["MetroE_Data"].range('A1').options(index=False, header=True).value = metroE

# delCw(mainwb.sheets["HB_Data"])
# delCw(mainwb.sheets["SMB_Data"])
# delCw(mainwb.sheets["MetroE_Data"])
# delCw(mainwb.sheets["MDU_HFC_Data"])
# delCw(mainwb.sheets["MDU_FTTP_Data"])
# delCw(mainwb.sheets["SDU-HFC_Data"])
# delCw(mainwb.sheets["SDU-HFC_Data"])
# delCw(mainwb.sheets["Brownfield_Data"])
# delCw(mainwb.sheets["SEO_Data"])
# delCw(mainwb.sheets["RBB_Data"])


Funnel = "blob:https://data.comcast.com/4b202101-2780-4218-a638-961f1ab2b6d9"
sphere = "https://sphere.comcast.com/g/api/scoop/6a708e19-fef3-497f-b007-dc63054c7be1/download"

=LET(existing,$A$3:INDEX($A:$A,ROW()-1),ids,VSTACK(IRR_Resi_Table[Request/Job: P2ID],IRR_Com_Table[Request/Job: P2ID])UNIQUE(FILTER(ids,(ids<>"")*ISNA(MATCH(ids,existing,0)),"")))